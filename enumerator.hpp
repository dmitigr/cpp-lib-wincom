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

#include "object.hpp"

namespace dmitigr::wincom {

class Enumerator final : public Unknown_api<Enumerator, IEnumVARIANT> {
  using Ua = Unknown_api<Enumerator, IEnumVARIANT>;
public:
  using Ua::Ua;

  explicit Enumerator(IUnknown* const api)
  {
    auto instance = query(api);
    swap(instance);
  }

  Enumerator clone() const
  {
    IEnumVARIANT* instance{};
    const auto err = detail::api(*this).Clone(&instance);
    DMITIGR_WINCOM_THROW_IF_ERROR(err, "cannot clone Enumerator");
    return Enumerator{instance};
  }
};

} // namespace dmitigr::wincom
