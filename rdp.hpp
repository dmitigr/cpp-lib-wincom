// -*- C++ -*-
//
// Copyright 2023 Dmitry Igrishin
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

#pragma once

#include "../base/noncopymove.hpp"
#include "../winbase/windows.hpp"
#include "exceptions.hpp"

#include <comdef.h> // avoid LNK2019
#include <comutil.h>
#include <Objbase.h>
#include <oleauto.h>
#include <Rdpencomapi.h>

#include <algorithm>
#include <memory>
#include <string>
#include <type_traits>

namespace dmitigr::wincom::rdp {

template<class DerivedType, class ApiType>
class Unknown_api : private Noncopy {
public:
  using Api = ApiType;
  using Derived = DerivedType;

  static_assert(std::is_base_of_v<IUnknown, Api>);

  virtual ~Unknown_api()
  {
    if (api_) {
      api_->Release();
      api_ = nullptr;
    }
  }

  static Derived query(IUnknown* const unknown)
  {
    assert(unknown);
    Api* api{};
    unknown->QueryInterface(&api);
    assert(api);
    return Derived{api};
  }

protected:
  Unknown_api() = default;

  explicit Unknown_api(Api* const api)
    : api_{api}
  {
    assert(api);
  }

  Unknown_api(Unknown_api&& rhs) noexcept
    : api_{rhs.api_}
  {
    rhs.api_ = nullptr;
  }

  Unknown_api& operator=(Unknown_api&& rhs) noexcept
  {
    Unknown_api tmp{std::move(rhs)};
    swap(tmp);
    return *this;
  }

  void swap(Unknown_api& rhs) noexcept
  {
    using std::swap;
    swap(api_, rhs.api_);
  }

  Api* api() const noexcept
  {
    return api_;
  }

private:
  Api* api_{};
};

class Invitation final
  : public Unknown_api<Invitation, IRDPSRAPIInvitation> {
public:
  using Super = Unknown_api<Invitation, IRDPSRAPIInvitation>;
  using Api = Super::Api;

  Invitation() = default;

  explicit Invitation(IRDPSRAPIInvitation* const api)
    : Super{api}
  {}

  template<class String>
  String connection() const
  {
    check(api(), L"invalid RDP invitation instance used");

    BSTR value;
    api()->get_ConnectionString(&value);
    _bstr_t tmp{value, false}; // take ownership
    return String(tmp);
  }
};

class Invitation_manager final
  : public Unknown_api<Invitation_manager, IRDPSRAPIInvitationManager> {
public:
  using Super = Unknown_api<Invitation_manager, IRDPSRAPIInvitationManager>;
  using Api = Super::Api;

  Invitation_manager() = default;

  Invitation_manager(IRDPSRAPIInvitationManager* const api)
    : Super{api}
  {}

  Invitation create_invitation(const std::wstring& group,
    const std::wstring& password, const long limit)
  {
    check(api(), L"invalid RDP invitation manager instance used");

    IRDPSRAPIInvitation* invitation{};
    const auto err = api()->CreateInvitation(
      nullptr,
      _bstr_t{group.c_str()},
      _bstr_t{password.c_str()},
      limit,
      &invitation);
    if (err != S_OK)
      throw Win_error{L"cannot create IRDPSRAPIInvitation instance", err};

    return Invitation{invitation};
  }
};

class Attendee_manager final
  : public Unknown_api<Attendee_manager, IRDPSRAPIAttendeeManager> {
public:
  using Super = Unknown_api<Attendee_manager, IRDPSRAPIAttendeeManager>;
  using Api = Super::Api;

  Attendee_manager() = default;

  Attendee_manager(IRDPSRAPIAttendeeManager* const api)
    : Super{api}
  {}
};

class Attendee final
  : public Unknown_api<Attendee, IRDPSRAPIAttendee> {
public:
  using Super = Unknown_api<Attendee, IRDPSRAPIAttendee>;
  using Api = Super::Api;

  Attendee() = default;

  Attendee(IRDPSRAPIAttendee* const api)
    : Super{api}
  {}

  void set_control_level(const CTRL_LEVEL level)
  {
    check(api(), L"invalid RDP attendee instance used");

    api()->put_ControlLevel(level);
  }
};

class Session_properties final
  : public Unknown_api<Session_properties, IRDPSRAPISessionProperties> {
public:
  using Super = Unknown_api<Session_properties, IRDPSRAPISessionProperties>;
  using Api = Super::Api;

  Session_properties() = default;

  Session_properties(IRDPSRAPISessionProperties* const api)
    : Super{api}
  {}

  Session_properties&
  set_clipboard_redirect_enabled(const bool value)
  {
    check(api(), L"invalid RDP session properties instance used");

    VARIANT val{};
    VariantInit(&val);
    val.vt = VT_BOOL;
    val.boolVal = value ? VARIANT_TRUE : VARIANT_FALSE;
    api()->put_Property(_bstr_t{"EnableClipboardRedirect"}, val);
    return *this;
  }

  Session_properties&
  set_clipboard_redirect_callback(IRDPSRAPIClipboardUseEvents* const value)
  {
    check(api(), L"invalid RDP session properties instance used");

    VARIANT val{};
    VariantInit(&val);
    val.vt = VT_UNKNOWN;
    val.punkVal = value;
    api()->put_Property(_bstr_t{"SetClipboardRedirectCallback"}, val);
    return *this;
  }
};

// -----------------------------------------------------------------------------

class Event_dispatcher : public _IRDPSessionEvents, private Noncopymove {
public:
  using Super = _IRDPSessionEvents;

  virtual void set_owner(void* owner) = 0;

  // IUnknown overrides

  HRESULT QueryInterface(REFIID id, void** const object) override
  {
    if (!object)
      return E_POINTER;

    if (id == __uuidof(Super))
      *object = static_cast<Super*>(this);
    else if (id == __uuidof(IDispatch))
      *object = static_cast<IDispatch*>(this);
    else if (id == __uuidof(IUnknown))
      *object = static_cast<IUnknown*>(this);
    else {
      *object = nullptr;
      return E_NOINTERFACE;
    }

    AddRef();
    return S_OK;
  }

  ULONG AddRef() override
  {
    return ++ref_count_;
  }

  ULONG Release() override
  {
    return ref_count_ = std::max(--ref_count_, ULONG(0));
  }

  // IDispatch overrides

  HRESULT GetTypeInfoCount(UINT* const info) override
  {
    *info = 0;
    return S_OK;
  }

  HRESULT GetTypeInfo(const UINT info, const LCID cid,
    ITypeInfo** const tinfo) override
  {
    if (info)
      return DISP_E_BADINDEX;

    *tinfo = nullptr;
    return S_OK;
  }

  HRESULT GetIDsOfNames(REFIID riid, LPOLESTR* const names,
    const UINT name_count, const LCID cid, DISPID* const disp_id) override
  {
    if (!disp_id)
      return E_OUTOFMEMORY;

    for (UINT i{}; i < name_count; ++i)
      disp_id[i] = DISPID_UNKNOWN;
    return DISP_E_UNKNOWNNAME;
  }

private:
  ULONG ref_count_{};
};

template<class Object, class ObjectInterface>
class Basic_com_object final : private Noncopymove {
public:
  ~Basic_com_object()
  {
    assert(api_);
    api_->Release();
    api_ = nullptr;
  }

  Basic_com_object()
    : Basic_com_object{CLSCTX_INPROC_SERVER}
  {}

  explicit Basic_com_object(const DWORD context)
  {
    if (const auto err = CoCreateInstance(
        __uuidof(Object),
        nullptr, // not a part of an aggregate
        context,
        __uuidof(ObjectInterface),
        reinterpret_cast<LPVOID*>(&api_)); err != S_OK)
      throw Win_error{L"cannot create COM object", err};
    assert(api_);
  }

  const ObjectInterface& api() const noexcept
  {
    return *api_;
  }

  ObjectInterface& api() noexcept
  {
    return const_cast<ObjectInterface&>(
      static_cast<const Basic_com_object*>(this)->api());
  }

private:
  ObjectInterface* api_{};
};

template<class BasicComObject>
class Basic_rdp_peer : private Noncopymove {
public:
  virtual ~Basic_rdp_peer()
  {
    if (point_) {
      point_->Unadvise(event_connection_token_);
      point_->Release();
      point_ = nullptr;
    }

    if (point_container_) {
      point_container_->Release();
      point_container_ = nullptr;
    }
  }

  Basic_rdp_peer(std::unique_ptr<BasicComObject> co,
    std::unique_ptr<Event_dispatcher> ed)
    : com_{std::move(co)}
    , event_dispatcher_{std::move(ed)}
  {
    const wchar_t* errmsg{};
    if (!com_)
      errmsg = L"invalid COM object";
    else if (!event_dispatcher_)
      errmsg = L"invalid event dispatcher";
    else if (com_->api().QueryInterface(&point_container_) != S_OK)
      errmsg = L"cannot query container for event connection point";
    else if (point_container_->FindConnectionPoint(
        __uuidof(_IRDPSessionEvents), &point_) != S_OK)
      errmsg = L"cannot find event connection point";
    else if (point_->Advise(event_dispatcher_.get(), &event_connection_token_) != S_OK)
      errmsg = L"cannot get event connection token";

    if (errmsg)
      throw Exception{errmsg};

    event_dispatcher_->set_owner(this);
  }

  const BasicComObject& com() const noexcept
  {
    return *com_;
  }

  BasicComObject& com() noexcept
  {
    return const_cast<BasicComObject&>(static_cast<const Basic_rdp_peer*>(this)->com());
  }

private:
  std::unique_ptr<BasicComObject> com_;
  std::unique_ptr<Event_dispatcher> event_dispatcher_;
  DWORD event_connection_token_{};
  IConnectionPointContainer* point_container_{};
  IConnectionPoint* point_{};
};

using Viewer = Basic_com_object<RDPViewer, IRDPSRAPIViewer>;
using Sharer = Basic_com_object<RDPSession, IRDPSRAPISharingSession>;

using Client_base = Basic_rdp_peer<Viewer>;
using Server_base = Basic_rdp_peer<Sharer>;

// -----------------------------------------------------------------------------
// Server
// -----------------------------------------------------------------------------

class Server final : public Server_base {
public:
  using Super = Server_base;
  using Super::Super;

  void open()
  {
    if (!is_open_) {
      const auto err = com().api().Open();
      if (err != S_OK)
        throw Win_error{L"cannot open RDP server", err};
      is_open_ = true;
    }
  }

  void close()
  {
    if (is_open_) {
      const auto err = com().api().Close();
      if (err != S_OK)
        throw Win_error{L"cannot close RDP server", err};
      is_open_ = false;
    }
  }

  bool is_open() const noexcept
  {
    return is_open_;
  }

  Invitation_manager invitation_manager()
  {
    IRDPSRAPIInvitationManager* api{};
    com().api().get_Invitations(&api);
    assert(api);
    return Invitation_manager{api};
  }

  Attendee_manager attendee_manager()
  {
    IRDPSRAPIAttendeeManager* api{};
    com().api().get_Attendees(&api);
    assert(api);
    return Attendee_manager{api};
  }

  Session_properties session_properties()
  {
    IRDPSRAPISessionProperties* api{};
    com().api().get_Properties(&api);
    assert(api);
    return Session_properties{api};
  }

  void pause()
  {
    const auto err = com().api().Pause();
    if (err != S_OK)
      throw Win_error{L"cannot pause RDP server", err};
  }

  void resume()
  {
    const auto err = com().api().Resume();
    if (err != S_OK)
      throw Win_error{L"cannot resume RDP server", err};
  }

private:
  bool is_open_{};
};

// -----------------------------------------------------------------------------
// Client
// -----------------------------------------------------------------------------

class Client final : public Client_base {
public:
  using Super = Client_base;
  using Super::Super;

  ~Client() override
  {
    try {
      close();
    } catch (...) {}
  }

  template<class ConnStr, class NameStr, class PassStr>
  void open(const ConnStr& connection_string,
    const NameStr& name, const PassStr& password)
  {
    const auto err = com().api().Connect(
      _bstr_t{connection_string.c_str()},
      _bstr_t{name.c_str()},
      _bstr_t{password.c_str()});
    if (err != S_OK)
      throw Win_error{L"cannot open RDP client", err};
  }

  void close()
  {
    const auto err = com().api().Disconnect();
    if (err != S_OK)
      throw Win_error{L"cannot close RDP client", err};
  }

  void set_control_level(const CTRL_LEVEL level)
  {
    const auto err = com().api().RequestControl(level);
    if (err != S_OK)
      throw Win_error{L"cannot set control level of RDP client", err};
  }

  Session_properties session_properties()
  {
    IRDPSRAPISessionProperties* api{};
    com().api().get_Properties(&api);
    assert(api);
    return Session_properties{api};
  }
};

} // namespace dmitigr::wincom::rdp
