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
#pragma comment(lib, "oleaut32")
#pragma comment(lib, "wbemuuid")

#include "../base/noncopymove.hpp"
#include "exceptions.hpp"
#include "object.hpp"

#include <cstdint>
#include <string>

#include <oaidl.h> // VARIANT
#include <oleauto.h> // VariantClear
#include <Wbemidl.h>

namespace dmitigr::wincom::wmi {

class ClassObject final :
    public Unknown_api<ClassObject, IWbemClassObject> {
  using Ua = Unknown_api<ClassObject, IWbemClassObject>;
public:
  using Ua::Ua;

  struct Value final : dmitigr::Noncopymove {
    ~Value()
    {
      VariantClear(&data);
    }

    template<class String>
    String as_str() const
    {
      if (type != CIM_STRING)
        throw std::logic_error{"cannot get value of VARIANT as string"};
      return detail::str<String>(data.bstrVal);
    }

    std::string as_string() const
    {
      return as_str<std::string>();
    }

    std::wstring as_wstring() const
    {
      return as_str<std::wstring>();
    }

    std::int8_t as_int8() const
    {
      if (type != CIM_SINT8)
        throw std::logic_error{"cannot get value of VARIANT as int8"};
      return static_cast<std::int8_t>(data.cVal);
    }

    std::uint8_t as_uint8() const
    {
      if (type != CIM_UINT8)
        throw std::logic_error{"cannot get value of VARIANT as uint8"};
      return static_cast<std::uint8_t>(data.bVal);
    }

    std::int16_t as_int16() const
    {
      if (type != CIM_SINT16)
        throw std::logic_error{"cannot get value of VARIANT as int16"};
      return data.iVal;
    }

    std::uint16_t as_uint16() const
    {
      if (type != CIM_UINT16)
        throw std::logic_error{"cannot get value of VARIANT as uint16"};
      return data.uiVal;
    }

    std::int32_t as_int32() const
    {
      if (type != CIM_SINT32)
        throw std::logic_error{"cannot get value of VARIANT as int32"};
      return data.intVal;
    }

    std::uint32_t as_uint32() const
    {
      if (type != CIM_UINT32)
        throw std::logic_error{"cannot get value of VARIANT as uint32"};
      return data.uintVal;
    }

    std::int64_t as_int64() const
    {
      if (type != CIM_SINT64)
        throw std::logic_error{"cannot get value of VARIANT as int64"};
      return data.llVal;
    }

    std::uint64_t as_uint64() const
    {
      if (type != CIM_UINT64)
        throw std::logic_error{"cannot get value of VARIANT as uint64"};
      return data.ullVal;
    }

    bool as_bool() const
    {
      if (type != CIM_BOOLEAN)
        throw std::logic_error{"cannot get value of VARIANT as bool"};
      return data.boolVal == VARIANT_TRUE;
    }

    DATE as_date() const
    {
      if (type != CIM_DATETIME)
        throw std::logic_error{"cannot get value of VARIANT as DATE"};
      return data.date;
    }

    VARIANT data{};
    CIMTYPE type{};
    long flavor{};

  private:
    friend ClassObject;
    Value(const VARIANT dat, const CIMTYPE tp, const long flv)
      : data{dat}
      , type{tp}
      , flavor{flv}
    {}
  };

  Value value(const LPCWSTR name) const
  {
    VARIANT data{};
    CIMTYPE type{};
    long flavor{};
    const auto err = detail::api(*this).Get(name, 0, &data, &type, &flavor);
    throw_if_error(err, "cannot get value of IWbemClassObject");
    return Value{data, type, flavor};
  }
};

class EnumClassObject final :
    public Unknown_api<EnumClassObject, IEnumWbemClassObject> {
  using Ua = Unknown_api<EnumClassObject, IEnumWbemClassObject>;
public:
  using Ua::Ua;

  ClassObject next(const long timeout = WBEM_INFINITE)
  {
    IWbemClassObject* result{};
    ULONG result_count{};
    const auto err = api().Next(timeout, 1, &result, &result_count);
    if (err == WBEM_S_NO_ERROR)
      return ClassObject{result};
    else if (err != WBEM_S_FALSE)
      throw_if_error(err, "cannot get next object of IEnumWbemClassObject");
    return ClassObject{};
  }
};

class Services final : public Unknown_api<Services, IWbemServices> {
  using Ua = Unknown_api<Services, IWbemServices>;
public:
  using Ua::Ua;

  template<class String>
  EnumClassObject exec_query(const String& query,
    const long flags = WBEM_FLAG_RETURN_IMMEDIATELY|WBEM_FLAG_FORWARD_ONLY,
    IWbemContext* const ctx = {}) const
  {
    IEnumWbemClassObject* result{};
    const auto err = detail::api(*this).ExecQuery(detail::bstr("WQL"),
      detail::bstr(query),
      flags,
      ctx,
      &result);
    throw_if_error(err, "cannot execute query to retrieve objects from"
      " WMI services");
    return EnumClassObject{result};
  }
};

class Locator final : public Unknown_api<Locator, IWbemLocator> {
  using Ua = Unknown_api<Locator, IWbemLocator>;
public:
  using Ua::Ua;

  Services connect_server(const BSTR network_resource,
    const BSTR user,
    const BSTR password,
    const BSTR locale,
    const long security_flags,
    const BSTR authority,
    IWbemContext* const ctx = {})
  {
    IWbemServices* result{};
    const auto err = api().ConnectServer(
      network_resource,
      user,
      password,
      locale,
      security_flags,
      authority,
      ctx,
      &result);
    throw_if_error(err, "cannot connect to "+
      detail::str<std::string>(network_resource));
    return Services{result};
  }
};

} // namespace dmitigr::wincom::wmi
