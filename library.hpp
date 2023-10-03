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
#include "exceptions.hpp"

#include <Objbase.h>

namespace dmitigr::wincom {

class Library final : private Noncopymove {
public:
  ~Library()
  {
    CoUninitialize();
  }

  explicit Library(const DWORD concurrency_model = COINIT_MULTITHREADED)
  {
    const auto err = CoInitializeEx(nullptr, concurrency_model);
    if (err != S_OK && err != S_FALSE)
      throw Win_error{L"cannot initialize COM library", err};
  }
};

} // namespace dmitigr::wincom
