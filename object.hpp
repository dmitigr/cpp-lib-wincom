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
#pragma comment(lib, "ole32")

#include "../base/noncopymove.hpp"
#include "exceptions.hpp"

#include <comdef.h> // avoid LNK2019
#include <Objbase.h>
#include <unknwn.h>
#include <WTypes.h> // MSHCTX, MSHLFLAGS

#include <algorithm>
#include <stdexcept>
#include <string>
#include <typeinfo>
#include <type_traits>

namespace dmitigr::wincom {

template<class DerivedType, class ApiType>
class Unknown_api {
public:
  using Api = ApiType;
  using Derived = DerivedType;

  static_assert(std::is_base_of_v<IUnknown, Api>);

  template<class U>
  static Derived query(U* const unknown)
  {
    static_assert(std::is_base_of_v<IUnknown, U>);

    static const std::string msg{
      "cannot obtain interface "
      + std::string{typeid(Api).name()}
      + " from "
      + std::string{typeid(U).name()}
      + " to make "
      + std::string{typeid(Derived).name()}};

    if (!unknown)
      throw std::invalid_argument{msg+": null input pointer"};
    Api* api{};
    unknown->QueryInterface(&api);
    if (!api)
      throw std::runtime_error{msg};
    return Derived{api};
  }

  virtual ~Unknown_api()
  {
    if (api_) {
      api_->Release();
      api_ = nullptr;
    }
  }

  Unknown_api() = default;

  explicit Unknown_api(Api* const api)
    : api_{api}
  {}

  Unknown_api(Unknown_api&& rhs) noexcept
    : api_{rhs.api_}
  {
    rhs.api_ = nullptr;
  }

  Unknown_api(const Unknown_api& rhs) noexcept
    : api_{rhs.api_}
  {
    if (api_)
      api_->AddRef();
  }

  Unknown_api& operator=(Unknown_api&& rhs) noexcept
  {
    Unknown_api tmp{std::move(rhs)};
    swap(tmp);
    return *this;
  }

  Unknown_api& operator=(const Unknown_api& rhs) noexcept
  {
    Unknown_api tmp{rhs};
    swap(tmp);
    return *this;
  }

  void swap(Unknown_api& rhs) noexcept
  {
    using std::swap;
    swap(api_, rhs.api_);
  }

  const Api& api() const
  {
    check(api_, "invalid "+std::string{typeid(Derived).name()}+" instance used");
    return *api_;
  }

  Api& api()
  {
    return const_cast<Api&>(static_cast<const Unknown_api*>(this)->api());
  }

  explicit operator bool() const noexcept
  {
    return static_cast<bool>(api_);
  }

private:
  template<class> friend class Ptr;

  Api* api_{};
};

template<class A>
class Ptr final : public Unknown_api<Ptr<A>, A> {
  using Ua = Unknown_api<Ptr, A>;
public:
  using Ua::Ua;

  A* get() const noexcept
  {
    return Ua::api_;
  }

  A* operator->() const noexcept
  {
    return get();
  }

  template<class T>
  Ptr<T> to() const
  {
    return Ptr<T>::query(Ua::api_);
  }
};

// -----------------------------------------------------------------------------

class Standard_marshaler final : public
  Unknown_api<Standard_marshaler, IMarshal> {
  using Ua = Unknown_api<Standard_marshaler, IMarshal>;
public:
  Standard_marshaler(REFIID riid, const LPUNKNOWN unknown,
    const MSHCTX dest_ctx, const MSHLFLAGS flags)
  {
    IMarshal* instance{};
    const auto err = CoGetStandardMarshal(riid, unknown, dest_ctx, nullptr,
      flags, &instance);
    throw_if_error(err, "cannot get standard marshaler");
    Ua tmp{instance};
    swap(tmp);
  }
};

// -----------------------------------------------------------------------------

template<class Object, class ObjectInterface>
class Basic_com_object : private Noncopy {
public:
  using Api = ObjectInterface;
  using Instance = Object;

  virtual ~Basic_com_object()
  {
    if (api_) {
      api_->Release();
      api_ = nullptr;
    }
  }

  Basic_com_object()
    : Basic_com_object{CLSCTX_INPROC_SERVER, nullptr}
  {}

  explicit Basic_com_object(ObjectInterface* const api)
    : api_{api}
  {}

  explicit Basic_com_object(const DWORD context_mask,
    IUnknown* const aggregate = {})
  {
    if (const auto err = CoCreateInstance(
        __uuidof(Object),
        aggregate,
        context_mask,
        __uuidof(ObjectInterface),
        reinterpret_cast<LPVOID*>(&api_)); err != S_OK)
      throw Win_error{"cannot create COM object", err};
    if (!api_)
      throw std::logic_error{"invalid COM instance created"};
  }

  Basic_com_object(Basic_com_object&& rhs) noexcept
    : api_{rhs.api_}
  {
    rhs.api_ = nullptr;
  }

  Basic_com_object& operator=(Basic_com_object&& rhs) noexcept
  {
    Basic_com_object tmp{std::move(rhs)};
    swap(tmp);
    return *this;
  }

  void swap(Basic_com_object& rhs) noexcept
  {
    using std::swap;
    swap(api_, rhs.api_);
  }

  const ObjectInterface& api() const noexcept
  {
    check(api_,
      "invalid "+std::string{typeid(Basic_com_object).name()}+" instance used");
    return *api_;
  }

  ObjectInterface& api() noexcept
  {
    return const_cast<ObjectInterface&>(
      static_cast<const Basic_com_object*>(this)->api());
  }

  explicit operator bool() const noexcept
  {
    return static_cast<bool>(api_);
  }

private:
  ObjectInterface* api_{};
};

// -----------------------------------------------------------------------------

namespace detail {

template<class ComObject>
[[nodiscard]] auto& api(const ComObject& com) noexcept
{
  using Com = std::decay_t<decltype(com)>;
  return const_cast<Com&>(com).api();
}

template<class String>
String str(BSTR value)
{
  _bstr_t tmp{value, false}; // take ownership
  return String(tmp);
}

template<class String, class Wrapper, class Api>
String str(const Wrapper& wrapper, HRESULT(Api::* getter)(BSTR*))
{
  BSTR value;
  (detail::api(wrapper).*getter)(&value);
  return str<String>(value);
}

inline auto* c_str(const std::string& s)
{
  return s.c_str();
}

inline auto* c_str(const std::wstring& s)
{
  return s.c_str();
}

inline auto* c_str(const char*& s)
{
  return s;
}

inline auto* c_str(const wchar_t*& s)
{
  return s;
}

template<typename String>
inline _bstr_t bstr(const String& s)
{
  return _bstr_t{c_str(s)};
}

} // namespace detail

} // namespace dmitigr::wincom
