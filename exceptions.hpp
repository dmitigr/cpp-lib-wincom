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

#include <exception>
#include <string>

namespace dmitigr::wincom {

class Exception : public std::exception {
public:
  explicit Exception(std::wstring message)
    : message_{std::move(message)}
  {}

  const char* what() const noexcept override
  {
    return "Exception (see message())";
  }

  const std::wstring& message() const noexcept
  {
    return message_;
  }

private:
  std::wstring message_;
};

class Win_error : public Exception {
public:
  Win_error(std::wstring message, const long code)
    : Exception{message.append(L" (error ").append(std::to_wstring(code)).append(L")")}
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
inline void check(const T& condition, std::wstring message)
{
  if (!static_cast<bool>(condition))
    throw Exception{std::move(message)};
}

} // namespace dmitigr::wincom
