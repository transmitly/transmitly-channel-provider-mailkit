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

using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using MimeKit.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Transmitly.ChannelProvider;

namespace Transmitly.MailKit
{
	internal sealed class MailKitChannelProviderClient(MailKitConfigurationOptions optionObj) : ChannelProviderClient<IEmail>
	{
		private readonly MailKitConfigurationOptions _optionObj = Guard.AgainstNull(optionObj);

		public override async Task<IReadOnlyCollection<IDispatchResult?>> DispatchAsync(IEmail email, IDispatchCommunicationContext communicationContext, CancellationToken cancellationToken)
		{
			Guard.AgainstNull(email);
			Guard.AgainstNull(communicationContext);

			var msg = new MimeMessage
			{
				MessageId = MimeUtils.GenerateMessageId()
			};
			msg.From.Add(email.From.ToMailboxAddress());
			msg.To.AddRange(email.To!.Select(m => m.ToMailboxAddress()));
			msg.Subject = email.Subject;

			if (email.Bcc != null)
				msg.Bcc.AddRange(email.Bcc.Select(x => x.ToMailboxAddress()));

			if (email.Cc != null)
				msg.Cc.AddRange(email.Cc.Select(x => x.ToMailboxAddress()));

			if (email.ReplyTo != null)
				msg.ReplyTo.AddRange(email.ReplyTo.Select(x => x.ToMailboxAddress()));

			var body = new BodyBuilder()
			{
				HtmlBody = email.HtmlBody,
				TextBody = email.TextBody
			};

			AddAttachments(body, email.Attachments, cancellationToken);

			msg.Body = body.ToMessageBody();

			var client = new SmtpClient();
			await Connect(client, cancellationToken).ConfigureAwait(false);
			string result = await Send(msg, client, cancellationToken).ConfigureAwait(false);
			var commResult = new MailKitSendResult
			{
				ResourceId = msg.MessageId,
				MessageString = result,
				DispatchStatus = DispatchStatus.Dispatched
			};
			SendDeliveryReports(email, communicationContext, commResult);
			return [commResult];
		}

		private void SendDeliveryReports(IEmail email, IDispatchCommunicationContext communicationContext, MailKitSendResult commResult)
		{
			switch (commResult.DispatchStatus)
			{
				case DispatchStatus.Exception:
					Error(communicationContext, email, [commResult]);
					break;
				case DispatchStatus.Dispatched:
					Dispatched(communicationContext, email, [commResult]);
					break;
			}
		}

		private static async Task<string> Send(MimeMessage msg, SmtpClient client, CancellationToken cancellationToken)
		{
			//If SmtpClient.Send() throws an exception, then sending the message failed. If it doesn't throw an exception, then it succeeded.
			//https://github.com/jstedfast/MailKit/issues/861#issuecomment-496497579
			var result = await client.SendAsync(msg, cancellationToken).ConfigureAwait(false);
			await client.DisconnectAsync(true, cancellationToken).ConfigureAwait(false);
			return result;
		}

		private async Task Connect(SmtpClient client, CancellationToken cancellationToken)
		{
			var secureSocketOption = _optionObj.UseSsl ? SecureSocketOptions.StartTls : SecureSocketOptions.Auto;
			await client.ConnectAsync(_optionObj.Host, _optionObj.Port ?? 0, secureSocketOption, cancellationToken).ConfigureAwait(false);
			await client.AuthenticateAsync(_optionObj.UserName, _optionObj.Password, cancellationToken).ConfigureAwait(false);
		}

		private static void AddAttachments(BodyBuilder body, IReadOnlyCollection<IAttachment> attachments, CancellationToken cancellationToken)
		{
			foreach (var attachment in attachments)
			{
				var (mediaType, subType) = GetContentType(attachment.ContentType);
				body.Attachments.Add(attachment.Name, attachment.ContentStream, new ContentType(mediaType, subType), cancellationToken);
			}
		}

		private static (string mediaType, string subType) GetContentType(string? contentType)
		{
			(string, string) unknownContentType = ("application", "octet-stream");
			if (string.IsNullOrWhiteSpace(contentType))
				return unknownContentType;
			var separateTypes = contentType!.Split('/');
			if (separateTypes.Length != 2)
				return unknownContentType;
			return (separateTypes[0], separateTypes[1]);
		}
	}
}