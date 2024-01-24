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
#include "exceptions.hpp"

#include <comdef.h> // avoid LNK2019
#include <Objbase.h>
#include <unknwn.h>

#include <algorithm>
#include <stdexcept>
#include <string>
#include <type_traits>

namespace dmitigr::wincom {

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
    if (!unknown)
      throw std::invalid_argument{"cannot query IUnknown: null pointer"};
    Api* api{};
    unknown->QueryInterface(&api);
    assert(api);
    return Derived{api};
  }

protected:
  Unknown_api() = default;

  explicit Unknown_api(Api* const api)
    : api_{api}
  {}

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

template<class Object, class ObjectInterface>
class Basic_com_object : private Noncopy {
public:
  virtual ~Basic_com_object()
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

namespace detail {

template<class Object, class ObjectInterface>
[[nodiscard]]
ObjectInterface& api(const Basic_com_object<Object, ObjectInterface>& com) noexcept
{
  using Com = std::decay_t<decltype(com)>;
  return const_cast<Com&>(com).api();
}

template<class String, class Wrapper, class Api>
String str(const Wrapper& wrapper, HRESULT(Api::* getter)())
{
  BSTR value;
  (detail::api(wrapper)->*getter)(&value);
  _bstr_t tmp{value, false}; // take ownership
  return String(tmp);
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
