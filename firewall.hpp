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
// #pragma comment(lib, "FirewallAPI")
// #pragma comment(lib, "Hnetcfg")

#include "object.hpp"

#include <netfw.h>

namespace dmitigr::wincom::firewall {

class Authorized_application final :
  public Basic_com_object<NetFwAuthorizedApplication, INetFwAuthorizedApplication> {
public:
  bool is_enabled() const
  {
    VARIANT_BOOL result{VARIANT_FALSE};
    detail::api(*this).get_Enabled(&result);
    return result == VARIANT_TRUE;
  }

  void set_enabled(const bool value)
  {
    const VARIANT_BOOL val{value ? VARIANT_TRUE : VARIANT_FALSE};
    api().put_Enabled(val);
  }

  NET_FW_IP_VERSION ip_version() const
  {
    NET_FW_IP_VERSION result{};
    detail::api(*this).get_IpVersion(&result);
    return result;
  }

  void set_ip_version(const NET_FW_IP_VERSION value)
  {
    api().put_IpVersion(value);
  }

  template<class String>
  String name() const
  {
    return get(&INetFwAuthorizedApplication::get_Name);
  }

  template<class String>
  void set_name(const String& value)
  {
    api().put_Name(_bstr_t{value.c_str()});
  }

  template<class String>
  String process_image_file_name() const
  {
    return get(&INetFwAuthorizedApplication::get_ProcessImageFileName);
  }

  template<class String>
  void set_process_image_file_name(const String& value)
  {
    api().put_ProcessImageFileName(_bstr_t{value.c_str()});
  }

private:
  template<class String>
  String get(HRESULT(INetFwAuthorizedApplication::* getter)()) const
  {
    BSTR value;
    (detail::api(*this)->*getter)(&value);
    _bstr_t tmp{value, false}; // take ownership
    return String(tmp);
  }
};

class Authorized_applications final :
  public Unknown_api<Authorized_applications, INetFwAuthorizedApplications> {
public:
  using Super = Unknown_api<Authorized_applications, INetFwAuthorizedApplications>;
  using Api = Super::Api;

  Authorized_applications() = default;

  explicit Authorized_applications(INetFwAuthorizedApplications* const api)
    : Super{api}
  {}

  void add(const Authorized_application& app)
  {
    check(api(), L"invalid firewall::Authorized_applications instance used");

    const auto err = api()->Add(&detail::api(app));
    if (err != S_OK)
      throw Win_error{L"cannot add application to firewall collection", err};
  }

  template<class String>
  void remove(const String& image_file_name)
  {
    check(api(), L"invalid firewall::Authorized_applications instance used");

    const auto err = api()->Remove(_bstr_t{image_file_name.c_str()});
    if (err != S_OK)
      throw Win_error{L"cannot add application to firewall collection", err};
  }
};

class Profile final : public Unknown_api<Profile, INetFwProfile> {
public:
  using Super = Unknown_api<Profile, INetFwProfile>;
  using Api = Super::Api;

  Profile() = default;

  explicit Profile(INetFwProfile* const api)
    : Super{api}
  {}

  Authorized_applications authorized_applications() const
  {
    check(api(), L"invalid firewall::Profile instance used");

    INetFwAuthorizedApplications* apps{};
    api()->get_AuthorizedApplications(&apps);
    return Authorized_applications{apps};
  }
};

class Policy final : public Unknown_api<Policy, INetFwPolicy> {
public:
  using Super = Unknown_api<Policy, INetFwPolicy>;
  using Api = Super::Api;

  Policy() = default;

  explicit Policy(INetFwPolicy* const api)
    : Super{api}
  {}

  Profile current_profile() const
  {
    check(api(), L"invalid firewall::Policy instance used");

    INetFwProfile* profile{};
    api()->get_CurrentProfile(&profile);
    return Profile{profile};
  }

  Profile profile(const NET_FW_PROFILE_TYPE value) const
  {
    INetFwProfile* profile{};
    api()->GetProfileByType(value, &profile);
    return Profile{profile};
  }
};

class Manager final : Basic_com_object<NetFwMgr, INetFwMgr> {
public:
  NET_FW_PROFILE_TYPE current_profile_type() const
  {
    NET_FW_PROFILE_TYPE result{};
    detail::api(*this).get_CurrentProfileType(&result);
    return result;
  }

  Policy local_policy() const
  {
    INetFwPolicy* policy{};
    detail::api(*this).get_LocalPolicy(&policy);
    return Policy{policy};
  }
};

} // namespace dmitigr::wincom::firewall
