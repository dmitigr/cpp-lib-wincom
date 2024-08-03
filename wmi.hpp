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
#include "../winbase/combase.hpp"
#include "exceptions.hpp"
#include "object.hpp"

#include <cstdint>
#include <stdexcept>
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

    Value()
    {
      VariantInit(&data);
      type = CIM_EMPTY;
    }

    Value(const VARIANT dat, const CIMTYPE tp, const long flv)
      : data{dat}
      , type{tp}
      , flavor{flv}
    {}

    std::string as_string_utf8() const
    {
      check(CIM_STRING, "UTF-8 string");
      return winbase::com::to_string(data.bstrVal);
    }

    std::string as_string_acp() const
    {
      check(CIM_STRING, "ACP string");
      return winbase::com::to_string(data.bstrVal, CP_ACP);
    }

    std::wstring as_wstring() const
    {
      check(CIM_STRING, "UTF-16 string");
      return winbase::com::to_wstring(data.bstrVal);
    }

    std::int8_t as_int8() const
    {
      check(CIM_SINT8, "int8");
      return static_cast<std::int8_t>(data.cVal);
    }

    std::uint8_t as_uint8() const
    {
      check(CIM_UINT8, "uint8");
      return static_cast<std::uint8_t>(data.bVal);
    }

    std::int16_t as_int16() const
    {
      check(CIM_SINT16, "int16");
      return data.iVal;
    }

    std::uint16_t as_uint16() const
    {
      check(CIM_UINT16, "uint16");
      return data.uiVal;
    }

    std::int32_t as_int32() const
    {
      check(CIM_SINT32, "int32");
      return data.intVal;
    }

    std::uint32_t as_uint32() const
    {
      check(CIM_UINT32, "uint32");
      return data.uintVal;
    }

    std::int64_t as_int64() const
    {
      check(CIM_SINT64, "int64");
      return data.llVal;
    }

    std::uint64_t as_uint64() const
    {
      check(CIM_UINT64, "uint64");
      return data.ullVal;
    }

    float as_real32() const
    {
      check(CIM_REAL32, "real32");
      return data.fltVal;
    }

    double as_real64() const
    {
      check(CIM_REAL32, "real64");
      return data.dblVal;
    }

    bool as_bool() const
    {
      check(CIM_BOOLEAN, "bool");
      return data.boolVal == VARIANT_TRUE;
    }

    DATE as_date() const
    {
      check(CIM_DATETIME, "DATE");
      return data.date;
    }

    VARIANT data{};
    CIMTYPE type{};
    long flavor{};

  private:
    void check(const CIMTYPE tp, const std::string& tpnm) const
    {
      if (type != tp)
        throw std::logic_error{"cannot get value of IWbemClassObject as "+tpnm};
    }
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

class Locator final : public Basic_com_object<WbemLocator, IWbemLocator> {
  using Bco = Basic_com_object<WbemLocator, IWbemLocator>;
public:
  using Bco::Bco;

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
    throw_if_error(err, "cannot connect to "
      +winbase::com::to_string(network_resource));
    return Services{result};
  }
};

} // namespace dmitigr::wincom::wmi
