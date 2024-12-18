# ⚠️ This repository has been archived. 
Development has moved to the [Transmitly SMTP Channel Provider Repository](https://github.com/transmitly/transmitly-channelprovider-smtp).
------------------------

## Transmitly.ChannelProvider.MailKit

A [Transmitly](https://github.com/transmitly/transmitly) channel provider that enables sending Email communications through SMTP using [MailKit](https://github.com/jstedfast/MailKit)

### Getting started

To use the MailKit channel provider, first install the [NuGet package](https://nuget.org/packages/transmitly.channelprovider.mailkit):

```shell
dotnet add package Transmitly.ChannelProvider.MailKit
```

Then add the channel provider using `AddMailKitSupport()`:

```csharp
using Transmitly;
//...
var communicationClient = new CommunicationsClientBuilder()
.AddMailKitSupport(options =>
{
	options.Host = "smtp.test.com";
	options.Port = 587;
	options.UseSsl = true;
	options.UserName = "Test";
	options.Password = "Password";
})
//Pipelines are the heart of Transmitly. Pipelines allow you to define your communications
//as a domain action. This allows your domain code to stay agnostic to the details of how you
//may send out a transactional communication.
.AddPipeline("first-pipeline", pipeline =>
{
	//AddEmail is a Channel that is core to the Transmitly library. 
	//AsAudienceAddress() is also a convience method that helps us create an audience address
	//Audience addresses can be anything, email, phone, or even a device/app Id for push notifications!
	pipeline.AddEmail("from@mydomain.com".AsAudienceAddress("My From Display Name"), email =>
	{
		//Transmitly is a bit different. All of our communication content is configured by templates.
		//Out of the box, we have static or string templates, file and even embedded template support.
		//There are multiple types of templates to get you started. You can even create templates 
		//specific to certain cultures!
		email.Subject.AddStringTemplate("Check out Transmit.ly!");
		email.HtmlBody.AddStringTemplate("Hey, check out this cool communciations library. <a href=\"https://transmit.ly\">")
		email.TextBody.AddStringTemplate("Hey, check out this cool communciations library. https://transmitly.ly");
	});
})
.BuildClient();

//Dispatch (send) the transsactional email to our friend Joe (joe@mydomain.com) using our configured SMTP server and our "first-pipeline" pipeline.
var result = await communicationsClient.DispatchAsync("first-pipeline", "joe@mydomain.com".AsAudienceAddress("Joe"), new { });
```
* Check out the [Transmitly](https://github.com/transmitly/transmitly) project for more details on what a channel provider is and how it can be used to improve how you manage your customer communications.


<picture>
  <source media="(prefers-color-scheme: dark)" srcset="https://github.com/transmitly/transmitly/assets/3877248/524f26c8-f670-4dfa-be78-badda0f48bfb">
  <img alt="an open-source project sponsored by CiLabs of Code Impressions, LLC" src="https://github.com/transmitly/transmitly/assets/3877248/34239edd-234d-4bee-9352-49d781716364" width="350" align="right">
</picture> 

---------------------------------------------------

_Copyright &copy; Code Impressions, LLC - Provided under the [Apache License, Version 2.0](http://apache.org/licenses/LICENSE-2.0.html)._
