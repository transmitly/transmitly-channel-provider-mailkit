// ﻿﻿Copyright (c) Code Impressions, LLC. All Rights Reserved.
//  
//  Licensed under the Apache License, Version 2.0 (the "License")
//  you may not use this file except in compliance with the License.
//  You may obtain a copy of the License at
//  
//      http://www.apache.org/licenses/LICENSE-2.0
//  
//  Unless required by applicable law or agreed to in writing, software
//  distributed under the License is distributed on an "AS IS" BASIS,
//  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//  See the License for the specific language governing permissions and
//  limitations under the License.

using System;
using System.Collections.Generic;

namespace Transmitly.MailKit
{
	internal sealed class MailKitSendResult : IDispatchResult
	{
		public string? ResourceId { get; set; }

		public IList<Exception> Exceptions { get; } = new List<Exception>();

		public IList<string> Messages { get; } = new List<string>();

		public string? MessageString { get; set; }

		public string? ChannelProviderId { get; set; }
		public string? ChannelId { get; set; }

		public DispatchStatus DispatchStatus { get; internal set; }

		public Exception? Exception { get; internal set; }
	}
}
