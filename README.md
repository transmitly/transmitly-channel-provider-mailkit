# Transmitly.ChannelProvider.MailKit

A [Transmitly](https://github.com/transmitly/transmitly) channel provider that enables sending Email communications through SMTP using [MailKit](https://github.com/jstedfast/MailKit)

### Getting started

To use the MailKit channel provider, first install the [NuGet package](https://nuget.org/packages/transmitly.channelprovider.mailkit):

```shell
dotnet add package Transmitly.ChannelProvider.MailKit
```

Then add the channel provider using `AddMailKitSupport()`:

```csharp
using Transmitly;
...
var communicationClient = new CommunicationsClientBuilder()
	.AddMailKitSupport(options =>
	{
		options.Host = "smtp.test.com";
		options.Port = 587;
		options.UseSsl = true;
		options.UserName = "Test";
		options.Password = "Password";
	});
```
* Check out the [Transmitly](https://github.com/transmitly/transmitly) project for more details on what a channel provider is and how it can be used to improve how you manage your customer communications.


<picture>
  <source media="(prefers-color-scheme: dark)" srcset="https://github.com/transmitly/transmitly/assets/3877248/524f26c8-f670-4dfa-be78-badda0f48bfb">
  <img alt="an open-source project sponsored by CiLabs of Code Impressions, LLC" src="https://github.com/transmitly/transmitly/assets/3877248/34239edd-234d-4bee-9352-49d781716364" width="350" align="right">
</picture> 

---------------------------------------------------

_Copyright &copy; Code Impressions, LLC - Provided under the [Apache License, Version 2.0](http://apache.org/licenses/LICENSE-2.0.html)._
