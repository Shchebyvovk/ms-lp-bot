require('dotenv').config();
const restify = require('restify');
const { BotFrameworkAdapter, ActivityTypes } = require('botbuilder');
const OpenAI = require('openai');

// ── OpenAI ────────────────────────────────────────────────────────────────────
const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

// Зберігаємо threadId для кожної розмови
// conversationId (LP/Bot) → OpenAI thread_id
const threads = new Map();

async function askAssistant(conversationId, userMessage) {
  // Якщо треду нема — створюємо новий
  if (!threads.has(conversationId)) {
    const thread = await openai.beta.threads.create();
    threads.set(conversationId, thread.id);
    console.log(`[OpenAI] New thread ${thread.id} for conv ${conversationId}`);
  }

  const threadId = threads.get(conversationId);

  // Додаємо повідомлення юзера в тред
  await openai.beta.threads.messages.create(threadId, {
    role: 'user',
    content: userMessage,
  });

  // Запускаємо асистента
  const run = await openai.beta.threads.runs.createAndPoll(threadId, {
    assistant_id: process.env.OPENAI_ASSISTANT_ID,
  });

  if (run.status !== 'completed') {
    console.error('[OpenAI] Run status:', run.status);
    return 'Вибачте, сталася помилка. Спробуйте ще раз.';
  }

  // Отримуємо останню відповідь асистента
  const messages = await openai.beta.threads.messages.list(threadId, {
    order: 'desc',
    limit: 1,
  });

  const reply = messages.data[0]?.content[0]?.text?.value;
  return reply || 'Немає відповіді від асистента.';
}

// ── Bot Framework Adapter ─────────────────────────────────────────────────────
const adapter = new BotFrameworkAdapter({
  appId: process.env.MicrosoftAppId,
  appPassword: process.env.MicrosoftAppPassword,
  channelAuthTenant: process.env.MicrosoftAppTenantId,
});

adapter.onTurnError = async (context, error) => {
  console.error('[onTurnError]', error);
  await context.sendActivity('Сталася помилка. Спробуйте ще раз.');
};

// ── Обробник повідомлень ──────────────────────────────────────────────────────
async function handleTurn(context) {
  const activity = context.activity;

  // Старт розмови — привітання
  if (
    activity.type === ActivityTypes.Event &&
    activity.name === 'CONVERSATION_START'
  ) {
    await context.sendActivity('Привіт! Чим можу допомогти?');
    return;
  }

  // Звичайне повідомлення від юзера
  if (activity.type === ActivityTypes.Message) {
    const userText = activity.text?.trim();
    if (!userText) return;

    const conversationId = activity.conversation?.id || 'default';

    console.log(`[MSG] conv=${conversationId} text="${userText}"`);

    // Відправляємо до OpenAI Assistant і чекаємо відповідь
    const reply = await askAssistant(conversationId, userText);

    await context.sendActivity(reply);
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