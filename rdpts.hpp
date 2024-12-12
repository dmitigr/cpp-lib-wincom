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

#include <chrono>
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

  enum class Network_connection_type : UINT {
    /// 56 Kbps.
    modem = 1,
    /// 256 Kbps to 2 Mbps.
    low = 2,
    /// 2 Mbps to 16 Mbps, with high latency.
    satellite = 3,
    /// 2 Mbps to 10 Mbps.
    broadband_high = 4,
    /// 10 Mbps or higher, with high latency.
    wan = 5,
    /// 10 Mbps or higher.
    lan = 6
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

  void set_overall_connection_timeout(const std::chrono::seconds value)
  {
    const auto err = api().put_overallConnectionTimeout(value.count());
    throw_if_error(err, "cannot set overall connection timeout");
  }

  std::chrono::seconds overall_connection_timeout() const
  {
    LONG result{};
    const auto err = detail::api(*this).get_overallConnectionTimeout(&result);
    throw_if_error(err, "cannot get overall connection timeout");
    return std::chrono::seconds{result};
  }

  void set_single_connection_timeout(const std::chrono::seconds value)
  {
    const auto err = api().put_singleConnectionTimeout(value.count());
    throw_if_error(err, "cannot set single connection timeout");
  }

  std::chrono::seconds single_connection_timeout() const
  {
    LONG result{};
    const auto err = detail::api(*this).get_singleConnectionTimeout(&result);
    throw_if_error(err, "cannot get single connection timeout");
    return std::chrono::seconds{result};
  }

  void set_shutdown_timeout(const std::chrono::seconds value)
  {
    const auto err = api().put_shutdownTimeout(value.count());
    throw_if_error(err, "cannot set shutdown timeout");
  }

  std::chrono::seconds shutdown_timeout() const
  {
    LONG result{};
    const auto err = detail::api(*this).get_shutdownTimeout(&result);
    throw_if_error(err, "cannot get shutdown timeout");
    return std::chrono::seconds{result};
  }

  void set_idle_timeout(const std::chrono::minutes value)
  {
    const auto err = api().put_MinutesToIdleTimeout(value.count());
    throw_if_error(err, "cannot set idle timeout");
  }

  std::chrono::minutes idle_timeout() const
  {
    LONG result{};
    const auto err = detail::api(*this).get_MinutesToIdleTimeout(&result);
    throw_if_error(err, "cannot get idle timeout");
    return std::chrono::minutes{result};
  }

  /// @param value The minimum valid value is `10000`.
  void set_keep_alive_interval(const std::chrono::milliseconds value)
  {
    const auto err = api().put_keepAliveInterval(value.count());
    throw_if_error(err, "cannot set keep-alive interval");
  }

  std::chrono::milliseconds keep_alive_interval() const
  {
    LONG result{};
    const auto err = detail::api(*this).get_keepAliveInterval(&result);
    throw_if_error(err, "cannot get keep-alive interval");
    return std::chrono::milliseconds{result};
  }

  // IMsRdpClientAdvancedSettings2

  void set_auto_reconnect_enabled(const bool value)
  {
    const VARIANT_BOOL val{value ? VARIANT_TRUE : VARIANT_FALSE};
    const auto err = api().put_EnableAutoReconnect(val);
    throw_if_error(err, "cannot set auto reconnect enabled");
  }

  bool is_auto_reconnect_enabled() const
  {
    VARIANT_BOOL result{VARIANT_FALSE};
    const auto err = detail::api(*this).get_EnableAutoReconnect(&result);
    throw_if_error(err, "cannot get auto reconnect enabled");
    return result == VARIANT_TRUE;
  }

  void set_max_reconnect_attempts(const LONG value)
  {
    const auto err = detail::api(*this).put_MaxReconnectAttempts(value);
    throw_if_error(err, "cannot set max reconnect attempts");
  }

  LONG max_reconnect_attempts() const
  {
    LONG result{};
    const auto err = detail::api(*this).get_MaxReconnectAttempts(&result);
    throw_if_error(err, "cannot get max reconnect attempts");
    return result;
  }

  // IMsRdpClientAdvancedSettings4

  void set_authentication_level(const Server_authentication value)
  {
    const auto err = api().put_AuthenticationLevel(
      static_cast<std::underlying_type_t<Server_authentication>>(value));
    throw_if_error(err, "cannot set authentication level");
  }

  void set(const Server_authentication value)
  {
    set_authentication_level(value);
  }

  Server_authentication authentication_level() const
  {
    UINT result{};
    detail::api(*this).get_AuthenticationLevel(&result);
    return Server_authentication{result};
  }

  // IMsRdpClientAdvancedSettings7

  void set_network_connection_type(const Network_connection_type value)
  {
    const auto err = api().put_NetworkConnectionType(
      static_cast<std::underlying_type_t<Network_connection_type>>(value));
    throw_if_error(err, "cannot set network connection type");
  }

  void set(const Network_connection_type value)
  {
    set_network_connection_type(value);
  }

  Network_connection_type network_connection_type() const
  {
    UINT result{};
    detail::api(*this).get_NetworkConnectionType(&result);
    return Network_connection_type{result};
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

  // IMsRdpClient7

  template<class String>
  String status_text(const UINT status_code) const
  {
    return String(detail::api(*this).GetStatusText(status_code));
  }

  // IMsRdpExtendedSettings

  void set_property_disable_auto_reconnect_component(const bool value)
  {
    VARIANT val{};
    VariantInit(&val);
    val.vt = VT_BOOL;
    val.boolVal = value ? VARIANT_TRUE : VARIANT_FALSE;
    const auto err = api<MSTSCLib::IMsRdpExtendedSettings>()
      .put_Property(detail::bstr("DisableAutoReconnectComponent"), &val);
    throw_if_error(err, "cannot disable auto reconnect component");
  }

private:
  Advise_sink_connection<Event_dispatcher> sink_;
};

} // namespace dmitigr::wincom::rdpts
