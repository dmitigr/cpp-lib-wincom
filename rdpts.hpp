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

#import "libid:8C11EFA1-92C3-11D1-BC1E-00C04FA31489"
#include <mstscax.tlh>

namespace dmitigr::wincom::rdpts {

class Advanced_settings final
  : public Unknown_api<Advanced_settings, MSTSCLib::IMsRdpClientAdvancedSettings8> {
  using Ua = Unknown_api<Advanced_settings, MSTSCLib::IMsRdpClientAdvancedSettings8>;
public:
  using Ua::Ua;

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
    api().put_SmartSizing(val);
  }

  bool is_smart_sizing_enabled() const noexcept
  {
    VARIANT_BOOL result{VARIANT_FALSE};
    detail::api(*this).get_SmartSizing(&result);
    return result == VARIANT_TRUE;
  }
};

class Client final :
    public Basic_com_object<MSTSCLib::MsRdpClient12, MSTSCLib::IMsRdpClient10> {
  using Bco = Basic_com_object<MSTSCLib::MsRdpClient12, MSTSCLib::IMsRdpClient10>;
public:
  using Bco::Bco;

  ~Client()
  {
    try {
      disconnect();
    } catch (...) {}
  }

  Advanced_settings advanced_settings() const
  {
    MSTSCLib::IMsRdpClientAdvancedSettings8* result{};
    detail::api(*this).get_AdvancedSettings9(&result);
    return Advanced_settings{result};
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
};

} // namespace dmitigr::wincom::rdpts
