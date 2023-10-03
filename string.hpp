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

#include <Objbase.h>

namespace dmitigr::wincom {

class String final : private Noncopymove {
public:
  String() = default;

  ~String()
  {
    CoTaskMemFree(value_);
  }

  explicit String(LPOLESTR value)
    : value_{value}
  {}

  LPOLESTR value() noexcept
  {
    return value_;
  }

  const LPOLESTR value() const noexcept
  {
    return value_;
  }

  bool is_valid() const noexcept
  {
    return static_cast<bool>(value_);
  }

  explicit operator bool() const noexcept
  {
    return is_valid();
  }

private:
  LPOLESTR value_{};
};

inline String to_com_string(REFCLSID id)
{
  LPOLESTR str{};
  if (StringFromCLSID(id, &str) == S_OK)
    return String{str};
  return String{};
}

} // namespace dmitigr::wincom
