// -*- C++ -*-
//
// Copyright 2024 Dmitry Igrishin
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

#include "../winbase/windows.hpp"
#include "exceptions.hpp"
#include "object.hpp"

#include <type_traits>

#import "libid:8C11EFA1-92C3-11D1-BC1E-00C04FA31489"
#include <mstscax.tlh>

namespace dmitigr::wincom::rdpts {

// -----------------------------------------------------------------------------
// Advanced_settings
// -----------------------------------------------------------------------------

class Advanced_settings final
  : public Unknown_api<Advanced_settings, MSTSCLib::IMsRdpClientAdvancedSettings8> {
  using Ua = Unknown_api<Advanced_settings, MSTSCLib::IMsRdpClientAdvancedSettings8>;
public:
  using Ua::Ua;

  enum class Server_authentication : UINT {
    /// No authentication of the server.
    disabled = 0,
    /// Server authentication is required.
    required = 1,
    /// Attempt authentication of the server.
    prompted = 2
  };

  // IMsRdpClientAdvancedSettings

  void set_rdp_port(const LONG value)
  {
    const auto err = api().put_RDPPort(value);
    throw_if_error(err, "cannot set RDP port");
  }

  LONG rdp_port() const
  {
    LONG result{};
    detail::api(*this).get_RDPPort(&result);
    return result;
  }

  void set_smart_sizing_enabled(const bool value)
  {
    const VARIANT_BOOL val{value ? VARIANT_TRUE : VARIANT_FALSE};
    const auto err = api().put_SmartSizing(val);
    throw_if_error(err, "cannot set smart sizing enabled");
  }

  bool is_smart_sizing_enabled() const noexcept
  {
    VARIANT_BOOL result{VARIANT_FALSE};
    detail::api(*this).get_SmartSizing(&result);
    return result == VARIANT_TRUE;
  }

  // IMsRdpClientAdvancedSettings4

  void set_authentication_level(const Server_authentication value)
  {
    const auto err = api().put_AuthenticationLevel(
      static_cast<std::underlying_type_t<Server_authentication>>(value));
    throw_if_error(err, "cannot set authentication level");
  }

  Server_authentication authentication_level() const
  {
    UINT result{};
    detail::api(*this).get_AuthenticationLevel(&result);
    return Server_authentication{result};
  }
};

// -----------------------------------------------------------------------------
// Client
// -----------------------------------------------------------------------------

class Client_event_dispatcher : public Advise_sink<MSTSCLib::IMsTscAxEvents> {};

class Client final : public Basic_com_object<
  MSTSCLib::MsRdpClient11NotSafeForScripting, MSTSCLib::IMsRdpClient10> {
  using Bco = Basic_com_object<MSTSCLib::MsRdpClient11NotSafeForScripting,
    MSTSCLib::IMsRdpClient10>;
public:
  using Bco::Bco;

  using Event_dispatcher = Client_event_dispatcher;

  ~Client()
  {
    try {
      disconnect();
    } catch (...) {}
  }

  explicit Client(std::unique_ptr<Event_dispatcher> sink)
    : sink_{api(), std::move(sink), this}
  {
    /*
     * Closing a window in which this ActiveX object is hosted leads to
     * releasing it, which causes failure upon of calling api().Release()
     * from ~Bco(). Thus, api.AddRef() call is required here.
     */
    api().AddRef();
  }

  Advanced_settings advanced_settings() const
  {
    MSTSCLib::IMsRdpClientAdvancedSettings8* result{};
    detail::api(*this).get_AdvancedSettings9(&result);
    return Advanced_settings{result};
  }

  // MsTscAxNotSafeForScripting

  template<class String>
  String version() const
  {
    return detail::str<String>(*this, &Api::get_Version);
  }

  /**
   * @param value DNS name or IP address.
   *
   * @remarks Must be set before calling the connect().
   */
  template<class String>
  void set_server(const String& value)
  {
    const auto err = api().put_Server(detail::bstr(value));
    throw_if_error(err, "cannot set Server property of RDP client");
  }

  template<class String>
  String server() const
  {
    return detail::str<String>(*this, &Api::get_Server);
  }

  template<class String>
  void set_user_name(const String& value)
  {
    const auto err = api().put_UserName(detail::bstr(value));
    throw_if_error(err, "cannot set UserName property of RDP client");
  }

  template<class String>
  String user_name() const
  {
    return detail::str<String>(*this, &Api::get_UserName);
  }

  void set_prompt_for_credentials_enabled(const bool value)
  {
    const VARIANT_BOOL val{value ? VARIANT_TRUE : VARIANT_FALSE};
    const auto err = api<MSTSCLib::IMsRdpClientNonScriptable3>()
      .put_PromptForCredentials(val);
    throw_if_error(err, "cannot set PromptForCredentials property of RDP client");
  }

  bool is_prompt_for_credentials_enabled() const noexcept
  {
    VARIANT_BOOL result{VARIANT_FALSE};
    detail::unconst(api<MSTSCLib::IMsRdpClientNonScriptable3>())
      .get_PromptForCredentials(&result);
    return result == VARIANT_TRUE;
  }

  void set_prompt_for_credentials_on_client_enabled(const bool value)
  {
    const VARIANT_BOOL val{value ? VARIANT_TRUE : VARIANT_FALSE};
    const auto err = api<MSTSCLib::IMsRdpClientNonScriptable4>()
      .put_PromptForCredsOnClient(val);
    throw_if_error(err, "cannot set PromptForCredsOnClient property of RDP client");
  }

  bool is_prompt_for_credentials_on_client_enabled() const noexcept
  {
    VARIANT_BOOL result{VARIANT_FALSE};
    detail::unconst(api<MSTSCLib::IMsRdpClientNonScriptable4>())
      .get_PromptForCredsOnClient(&result);
    return result == VARIANT_TRUE;
  }

  void set_desktop_height(const LONG value)
  {
    const auto err = api().put_DesktopHeight(value);
    throw_if_error(err, "cannot set DesktopHeight property of RDP client");
  }

  LONG desktop_height() const
  {
    LONG result{};
    detail::api(*this).get_DesktopHeight(&result);
    return result;
  }

  void set_desktop_width(const LONG value)
  {
    const auto err = api().put_DesktopWidth(value);
    throw_if_error(err, "cannot set DesktopWidth property of RDP client");
  }

  LONG desktop_width() const
  {
    LONG result{};
    detail::api(*this).get_DesktopWidth(&result);
    return result;
  }

  short connection_state() const
  {
    short result{};
    detail::api(*this).get_Connected(&result);
    return result;
  }

  void connect()
  {
    const auto err = api().Connect();
    throw_if_error(err, "cannot initiate connection to remote RDP server");
  }

  void disconnect()
  {
    const auto err = api().Disconnect();
    throw_if_error(err, "cannot disconnect from remote RDP server");
  }

  MSTSCLib::ControlReconnectStatus reconnect(const ULONG width, const ULONG height)
  {
    return api().Reconnect(width, height);
  }

  // MsRdpClient9NotSafeForScripting

  /**
   * @remarks This method:
   *   - will throw if not logged into the user session;
   *   - may throw until some amount of time have elapsed after logging into
   *   the user session.
   */
  void update_session_display_settings(const ULONG desktop_width,
    const ULONG desktop_height, const ULONG physical_width,
    const ULONG physical_height, const ULONG orientation,
    const ULONG desktop_scale_factor, const ULONG device_scale_factor)
  {
    const auto err = [&]
    {
      try {
        // UpdateSessionDisplaySettings() throws exception if no login complete.
        return api().UpdateSessionDisplaySettings(desktop_width,
          desktop_height, physical_width, physical_height, orientation,
          desktop_scale_factor, device_scale_factor);
      } catch (...) {
        throw std::runtime_error{"cannot update RDP session display settings"};
      }
    }();
    throw_if_error(err, "cannot update RDP session display settings");
  }

  void sync_session_display_settings()
  {
    const auto err = api().SyncSessionDisplaySettings();
    throw_if_error(err, "cannot synchronize RDP session display settings");
  }

private:
  Advise_sink_connection<Event_dispatcher> sink_;
};

} // namespace dmitigr::wincom::rdpts
