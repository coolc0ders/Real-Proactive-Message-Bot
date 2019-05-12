// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
//
// Generated with Bot Builder V4 SDK Template for Visual Studio EchoBot v4.3.0

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Security.Claims;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Connector;
using Microsoft.Bot.Connector.Authentication;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Schema.Teams;

namespace ProactiveTeamsBot.Bots
{
    public class ProactiveBot : ActivityHandler
    {
        double _secondsToReply = 3;
        ICredentialProvider _credentialProvider;

        public ProactiveBot(ICredentialProvider credentialProvider)
        {
            _credentialProvider = credentialProvider;
        }

        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            try
            {
                await turnContext.SendActivityAsync(MessageFactory.Text($"I'll reply to you in {_secondsToReply} seconds."));
                QueueReplyAndSendItProactively(turnContext).Wait();
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
                throw e;
            }
        }

        public async Task QueueReplyAndSendItProactively(ITurnContext turnContext)
        {
            string conversationMessage = "I created my own conversation.";
            string replyMessage = "I proactively replied to this conversation.";

            var task = Task.Run(async () =>
            {
                await Task.Delay(TimeSpan.FromSeconds(_secondsToReply));
                // Let the Bot Proactively create a a conversation.
                var response = await CreateConversation(conversationMessage, turnContext);

                // Reply to the conversation which the bot created.
                await ProactivelyReplyToConversation(response.Id, replyMessage, turnContext);

                return Task.CompletedTask;
            });
            await task;
        }

        public async Task<ConversationResourceResponse> CreateConversation(string message, ITurnContext turnContext )
        {
            //var teamsContext = turnContext.TurnState.Get<ITeamsContext>();
            ConnectorClient _client = new ConnectorClient(new Uri(turnContext.Activity.ServiceUrl), await GetMicrosoftAppCredentialsAsync(turnContext), new HttpClient());
            var channelData = turnContext.Activity.GetChannelData<TeamsChannelData>();

            var conversationParameter = new ConversationParameters
            {
                Bot = turnContext.Activity.Recipient,
                IsGroup = true,
                ChannelData = channelData,
                TenantId = channelData.Tenant.Id,
                Activity = MessageFactory.Text(message)
            };
            var response = await _client.Conversations.CreateConversationAsync(conversationParameter);
            return response;
        }

        public async Task ProactivelyReplyToConversation(string conversationId, string message, ITurnContext turnContext)
        {
            ConnectorClient _client = new ConnectorClient(new Uri(turnContext.Activity.ServiceUrl), await GetMicrosoftAppCredentialsAsync(turnContext), new HttpClient());
            var reply = MessageFactory.Text(message);
            reply.Conversation = new ConversationAccount(isGroup: true, id: conversationId);
            await _client.Conversations.SendToConversationAsync(reply);
        }

        private async Task<MicrosoftAppCredentials> GetMicrosoftAppCredentialsAsync(ITurnContext turnContext)
        {
            ClaimsIdentity claimsIdentity = turnContext.TurnState.Get<ClaimsIdentity>("BotIdentity");

            Claim botAppIdClaim = claimsIdentity.Claims?.SingleOrDefault(claim => claim.Type == AuthenticationConstants.AudienceClaim)
                ??
                claimsIdentity.Claims?.SingleOrDefault(claim => claim.Type == AuthenticationConstants.AppIdClaim);

            string appPassword = await _credentialProvider.GetAppPasswordAsync(botAppIdClaim.Value).ConfigureAwait(false);
            return new MicrosoftAppCredentials(botAppIdClaim.Value, appPassword);
        }

        protected override async Task OnMembersAddedAsync(IList<ChannelAccount> membersAdded, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            foreach (var member in membersAdded)
            {
                if (member.Id != turnContext.Activity.Recipient.Id)
                {
                    await turnContext.SendActivityAsync(MessageFactory.Text($"Hello and Welcome!"), cancellationToken);
                }
            }
        }
    }
}
