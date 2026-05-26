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
  await context.sendActivity('Вибачте, сталася помилка.');
};

// ── Transfer to Agent ─────────────────────────────────────────────────────────
async function transferToAgent(context, conversationId) {
  console.log('[TRANSFER] Sending handoff.initiate event');
  
  transferredConversations.add(conversationId);
  
  // Спочатку текстове повідомлення
  await context.sendActivity('Зʼєдную вас з оператором. Зачекайте, будь ласка...');
  
  // Потім handoff event
  await context.sendActivity({
    type: 'event',
    name: 'handoff.initiate',
    value: {
      skill: 'agent-after-ms-bot'
    }
  });
  
  console.log('[TRANSFER] Handoff event sent');
}

// ── Обробник повідомлень ──────────────────────────────────────────────────────
async function handleTurn(context) {
  const activity = context.activity;
  const conversationId = activity.conversation?.id || 'default';

  // Якщо розмова передана — ігноруємо всі повідомлення
  if (transferredConversations.has(conversationId)) {
    console.log(`[IGNORED] Message from transferred conversation: ${conversationId}`);
    return;
  }

  // Старт розмови
  if (
    activity.type === ActivityTypes.Event &&
    activity.name === 'CONVERSATION_START'
  ) {
    await context.sendActivity(
      'Привіт! Я віртуальний асистент. Чим можу допомогти?\n\n' +
      'Якщо потрібен оператор, напишіть: **оператор**'
    );
    return;
  }

  // Звичайне повідомлення
  if (activity.type === ActivityTypes.Message) {
    const userText = activity.text?.trim();
    if (!userText) return;

    console.log(`[MSG] conv=${conversationId} text="${userText}"`);

    // Перевіряємо чи це команда на трансфер
    const lowerText = userText.toLowerCase();
    if (
      lowerText === 'оператор' ||
      lowerText === 'агент' ||
      lowerText === 'людина' ||
      lowerText === 'live agent' ||
      lowerText === 'human'
    ) {
      await transferToAgent(context, conversationId);
      return;
    }

    // Інакше питаємо OpenAI
    const reply = await askAssistant(conversationId, userText);
    
    if (reply) {
      await context.sendActivity(reply);
    } else {
      await context.sendActivity('Вибачте, не зміг сформувати відповідь. Спробуйте ще раз.');
    }
  }
}

// ── HTTP сервер ───────────────────────────────────────────────────────────────
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