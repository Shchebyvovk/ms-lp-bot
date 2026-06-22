require('dotenv').config();
const restify = require('restify');
const { BotFrameworkAdapter, ActivityTypes } = require('botbuilder');
const OpenAI = require('openai');

// ── OpenAI ────────────────────────────────────────────────────────────────────
const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
const threads = new Map();
const transferredConversations = new Set();

async function askAssistant(conversationId, userMessage) {
  if (!threads.has(conversationId)) {
    const thread = await openai.beta.threads.create();
    threads.set(conversationId, thread.id);
    console.log(`[OpenAI] New thread ${thread.id} for conv ${conversationId}`);
  }

  const threadId = threads.get(conversationId);

  await openai.beta.threads.messages.create(threadId, {
    role: 'user',
    content: userMessage,
  });

  const run = await openai.beta.threads.runs.createAndPoll(threadId, {
    assistant_id: process.env.OPENAI_ASSISTANT_ID,
  });

  if (run.status !== 'completed') {
    console.error('[OpenAI] Run status:', run.status);
    return null;
  }

  const messages = await openai.beta.threads.messages.list(threadId, {
    order: 'desc',
    limit: 1,
  });

  const reply = messages.data[0]?.content[0]?.text?.value;
  return reply || null;
}

// ── Bot Framework Adapter ─────────────────────────────────────────────────────
const adapter = new BotFrameworkAdapter({
  appId: process.env.MicrosoftAppId,
  appPassword: process.env.MicrosoftAppPassword,
  channelAuthTenant: process.env.MicrosoftAppTenantId,
});

adapter.onTurnError = async (context, error) => {
  console.error('[onTurnError]', error);
  await context.sendActivity('Sorry, an error occurred. Please try again.');
};

// ── Transfer to Agent ─────────────────────────────────────────────────────────
async function transferToAgent(context, conversationId) {
  console.log('[TRANSFER] Initiating transfer to skill: agent-after-ms-bot');
  
  transferredConversations.add(conversationId);
  
  // Message for customer (English)
  await context.sendActivity('Connecting you to an agent. Please wait...');
  
  // Welcome message for agent
  await context.sendActivity({
    type: 'message',
    text: '👋 A customer has requested to speak with an agent. Feel free to take over the conversation.',
    channelData: {
      messageAudience: 'AGENTS_AND_MANAGERS'
    }
  });
  
  // Transfer action
  await context.sendActivity({
    type: 'message',
    channelData: {
      action: {
        name: 'TRANSFER',
        parameters: {
          skill: 'agent-after-ms-bot'
        }
      }
    }
  });
  
  console.log('[TRANSFER] Transfer initiated with agent greeting');
}

// ── Message Handler ───────────────────────────────────────────────────────────
async function handleTurn(context) {
  const activity = context.activity;
  const conversationId = activity.conversation?.id || 'default';

  // Ignore transferred conversations
  if (transferredConversations.has(conversationId)) {
    console.log(`[IGNORED] Transferred conversation: ${conversationId}`);
    return;
  }

  // Conversation start
  if (
    activity.type === ActivityTypes.Event &&
    activity.name === 'CONVERSATION_START'
  ) {
    await context.sendActivity(
      'Hello! I am a virtual assistant. How can I help you?\n\n' +
      'If you need to speak with an agent, type: **agent**'
    );
    return;
  }

  // Regular message
  if (activity.type === ActivityTypes.Message) {
    const userText = activity.text?.trim();
    if (!userText) return;

    console.log(`[MSG] conv=${conversationId} text="${userText}"`);

    // Transfer command
    const lowerText = userText.toLowerCase();
    if (
      lowerText === 'agent' ||
      lowerText === 'operator' ||
      lowerText === 'human' ||
      lowerText === 'live agent' ||
      lowerText === 'help' ||
      lowerText === 'support'
    ) {
      await transferToAgent(context, conversationId);
      return;
    }

    // ✅ SHOW TYPING INDICATOR
    console.log('[TYPING] Showing typing indicator...');
    await context.sendActivity({
      type: ActivityTypes.Typing
    });

    // OpenAI Assistant
    const reply = await askAssistant(conversationId, userText);
    
    if (reply) {
      await context.sendActivity(reply);
      console.log('[REPLY] Message sent to user');
    } else {
      await context.sendActivity('Sorry, I could not generate a response. Please try again.');
    }
  }
}

// ── HTTP Server ───────────────────────────────────────────────────────────────
const server = restify.createServer({ name: 'LP-Bot' });
server.use(restify.plugins.bodyParser());

server.post('/api/messages', async (req, res) => {
  await adapter.processActivity(req, res, handleTurn);
});

server.get('/health', (req, res, next) => {
  res.send(200, { status: 'ok' });
  return next();
});

const port = process.env.PORT || 3978;
server.listen(port, () => {
  console.log(`✅ Bot running on port ${port}`);
});
