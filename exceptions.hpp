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

#include <new>
#include <stdexcept>
#include <string>
#include <utility>

#include <Winerror.h>

namespace dmitigr::wincom {

/// FIXME: inherit from std::system_error
class Win_error final : public std::runtime_error {
public:
  Win_error(std::string message, const long code)
    : std::runtime_error{message
        .append(" (error ").append(std::to_string(code)).append(")")}
    , code_{code}
  {}

  long code() const noexcept
  {
    return code_;
  }

private:
  long code_{};
};

template<typename T>
inline void check(const T& condition, const std::string& message)
{
  if (!static_cast<bool>(condition))
    throw std::logic_error{message};
}

[[deprecated]]
inline void throw_if_error(const HRESULT err, std::string message)
{
  if (FAILED(err)) {
    if (err == E_OUTOFMEMORY)
      throw std::bad_alloc{};
    else
      throw Win_error{std::move(message), err};
  }
}

} // namespace dmitigr::wincom

#define DMITIGR_WINCOM_THROW_IF_ERROR(err, msg) \
  do {                                          \
    if (const HRESULT e{err}; FAILED(e)) {      \
      if (e == E_OUTOFMEMORY)                   \
        throw std::bad_alloc{};                 \
      else                                      \
        throw Win_error{std::move(msg), e};     \
    }                                           \
  } while (false)
