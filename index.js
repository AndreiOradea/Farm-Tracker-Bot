const {
  Client,
  GatewayIntentBits,
  Partials,
  SlashCommandBuilder,
  REST,
  Routes,
  ActionRowBuilder,
  ButtonBuilder,
  ButtonStyle,
  AttachmentBuilder,
  EmbedBuilder,
} = require('discord.js');
const ExcelJS = require('exceljs');
const express = require('express');
const fs = require('fs');
const path = require('path');

// ========== VARIABILE DE MEDIU ==========
const TOKEN = process.env.TOKEN;
const CLIENT_ID = process.env.CLIENT_ID;
const GUILD_ID = process.env.GUILD_ID || null;

// ========== CONFIG ==========
const STOP_TIMEOUT_MINUTES = 12;
const CHECK_INTERVAL_MS = 60 * 1000;
// ============================

const client = new Client({
  intents: [
    GatewayIntentBits.Guilds,
    GatewayIntentBits.GuildMessages,
    GatewayIntentBits.MessageContent,
  ],
  partials: [Partials.Channel],
});

const channelFarmState = new Map();

// ========== KEEP-ALIVE ==========
const app = express();
const PORT = process.env.PORT || 3000;

app.get('/', (req, res) => {
  res.send('Farm Bot is online ✅');
});

app.listen(PORT, () => {
  console.log(`Keep-alive server running on port ${PORT}`);
});

// ========== PARSER FLEXIBIL ==========
function extractFarmData(msg) {
  let text = msg.content || '';

  if (msg.embeds?.length > 0) {
    text += '\n' + msg.embeds.map(e => {
      const parts = [];
      if (e.title) parts.push(e.title);
      if (e.description) parts.push(e.description);
      if (e.fields) {
        e.fields.forEach(f => {
          parts.push(`${f.name}: ${f.value}`);
        });
      }
      if (e.footer?.text) parts.push(e.footer.text);
      return parts.join('\n');
    }).join('\n');
  }

  // Căutăm orice variantă de "coins" + număr
  // Acceptă: "Coins Collected: 50", "Coins Collected in 3 Minutes: 38", "50 coins" etc.
  const patterns = [
    /Coins Collected(?: in \d+ Minutes)?[:\s]*(\d+)/i,
    /Coins[:\s]*(\d+)/i,
    /(\d+)\s*coins?/i,
  ];

  let coins = null;
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) {
      coins = parseInt(match[1], 10);
      break;
    }
  }

  if (coins === null || isNaN(coins)) return null;

  // Încercăm să extragem și estimated / bag / role dacă există
  const estimatedMatch = text.match(/Estimated Coins per Hour[:\s]*(\d+)/i);
  const bagMatch = text.match(/Bag[:\s]*(.+)/i);
  const roleMatch = text.match(/(?:Your Role|Role)[:\s]*(\w+)/i);

  return {
    coins,
    estimated: estimatedMatch ? parseInt(estimatedMatch[1], 10) : null,
    bag: bagMatch ? bagMatch[1].trim() : (roleMatch ? roleMatch[1] : 'Unknown'),
    timestamp: msg.createdAt,
    id: msg.id,
    url: msg.url,
    rawText: text.slice(0, 200)
  };
}

// ========== FETCH ÎMBUNĂTĂȚIT ==========
async function fetchFarmMessages(channel, hours = null) {
  const limit = 100;
  let allMessages = [];
  let lastId = null;
  const maxFetch = 1500; // mai multe mesaje
  const cutoff = hours ? Date.now() - (hours * 60 * 60 * 1000) : null;

  while (allMessages.length < maxFetch) {
    const options = { limit };
    if (lastId) options.before = lastId;

    const messages = await channel.messages.fetch(options);
    if (messages.size === 0) break;

    const batch = [...messages.values()];
    allMessages.push(...batch);
    lastId = messages.last().id;

    // Oprim dacă am ajuns înainte de cutoff
    if (cutoff && messages.last().createdTimestamp < cutoff) {
      break;
    }
  }

  let farmMessages = allMessages
    .map(extractFarmData)
    .filter(d => d !== null);

  if (cutoff) {
    farmMessages = farmMessages.filter(m => m.timestamp.getTime() >= cutoff);
  }

  // Sortăm de la cel mai vechi la cel mai nou
  farmMessages.sort((a, b) => a.timestamp - b.timestamp);
  return farmMessages;
}

function formatTime(date) {
  return date.toLocaleString('ro-RO', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
  });
}

function groupIntoSessions(farmMessages, gapMinutes = 12) {
  if (farmMessages.length === 0) return [];

  const sessions = [];
  let current = [farmMessages[0]];

  for (let i = 1; i < farmMessages.length; i++) {
    const gap = (farmMessages[i].timestamp - farmMessages[i - 1].timestamp) / 60000;
    if (gap > gapMinutes) {
      sessions.push(current);
      current = [farmMessages[i]];
    } else {
      current.push(farmMessages[i]);
    }
  }
  sessions.push(current);
  return sessions;
}

function buildSummaryEmbed(farmMessages, hours = null) {
  if (farmMessages.length === 0) {
    return new EmbedBuilder()
      .setColor(0xff0000)
      .setTitle('❌ Niciun farm găsit')
      .setDescription(
        hours
          ? `Nu am găsit mesaje cu coins în ultimele **${hours} ore**.`
          : 'Nu am găsit mesaje cu coins în istoricul recent.\nÎncearcă `/farm now ore:6` sau `/farm history`.'
      );
  }

  const start = farmMessages[0];
  const end = farmMessages[farmMessages.length - 1];
  const totalCoins = farmMessages.reduce((sum, m) => sum + m.coins, 0);
  const durationMs = end.timestamp - start.timestamp;
  const durationMin = Math.max(1, Math.round(durationMs / 60000));
  const hoursPart = Math.floor(durationMin / 60);
  const minutesPart = durationMin % 60;
  const avgPerHour = Math.round((totalCoins / durationMin) * 60);

  return new EmbedBuilder()
    .setColor(0x00ff99)
    .setTitle('📊 Rezumat Farm')
    .addFields(
      { name: '🟢 Start', value: formatTime(start.timestamp), inline: true },
      { name: '🔴 Stop', value: formatTime(end.timestamp), inline: true },
      { name: '⏱️ Durată', value: `${hoursPart > 0 ? `${hoursPart}h ` : ''}${minutesPart} min`, inline: true },
      { name: '💰 Total Coins', value: `**${totalCoins.toLocaleString('ro-RO')}**`, inline: true },
      { name: '📈 Mesaje', value: `${farmMessages.length}`, inline: true },
      { name: '⚡ Medie / oră', value: `~${avgPerHour.toLocaleString('ro-RO')}`, inline: true }
    )
    .setFooter({ text: hours ? `Filtru: ultimele ${hours} ore` : 'Toate mesajele recente' })
    .setTimestamp();
}

function createButtons(hours = null) {
  return new ActionRowBuilder().addComponents(
    new ButtonBuilder()
      .setCustomId(`recalc_${hours || 'all'}`)
      .setLabel('🔄 Recalculează')
      .setStyle(ButtonStyle.Primary),
    new ButtonBuilder()
      .setCustomId(`export_csv_${hours || 'all'}`)
      .setLabel('📄 Export CSV')
      .setStyle(ButtonStyle.Secondary),
    new ButtonBuilder()
      .setCustomId(`export_excel_${hours || 'all'}`)
      .setLabel('📊 Export Excel')
      .setStyle(ButtonStyle.Success)
  );
}

async function generateCSV(farmMessages) {
  const header = 'Timestamp,Coins,Estimated,Bag/Role,Message URL\n';
  const rows = farmMessages.map(m =>
    `"${formatTime(m.timestamp)}",${m.coins},${m.estimated || ''},"${m.bag}","${m.url}"`
  ).join('\n');

  const filePath = path.join(__dirname, `farm_${Date.now()}.csv`);
  fs.writeFileSync(filePath, header + rows, 'utf8');
  return filePath;
}

async function generateExcel(farmMessages) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Farm Data');

  sheet.columns = [
    { header: 'Timestamp', key: 'timestamp', width: 22 },
    { header: 'Coins', key: 'coins', width: 12 },
    { header: 'Estimated / Hour', key: 'estimated', width: 18 },
    { header: 'Bag / Role', key: 'bag', width: 15 },
    { header: 'Message URL', key: 'url', width: 50 },
  ];

  farmMessages.forEach(m => {
    sheet.addRow({
      timestamp: formatTime(m.timestamp),
      coins: m.coins,
      estimated: m.estimated || '',
      bag: m.bag,
      url: m.url,
    });
  });

  sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  sheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF1F4E79' },
  };

  const filePath = path.join(__dirname, `farm_${Date.now()}.xlsx`);
  await workbook.xlsx.writeFile(filePath);
  return filePath;
}

// ========== EVENTS ==========

client.once('clientReady', async () => {
  console.log(`✅ Bot online ca ${client.user.tag}`);

  const commands = [
    new SlashCommandBuilder()
      .setName('farm')
      .setDescription('Tracker de coins / farm')
      .addSubcommand(sub =>
        sub.setName('now')
          .setDescription('Analizează farm-ul recent')
          .addIntegerOption(opt =>
            opt.setName('ore')
              .setDescription('Ultimele X ore (opțional)')
              .setMinValue(1)
              .setMaxValue(168)
          )
      )
      .addSubcommand(sub =>
        sub.setName('history')
          .setDescription('Arată ultimele sesiuni de farm')
          .addIntegerOption(opt =>
            opt.setName('numar')
              .setDescription('Câte sesiuni (implicit 5)')
              .setMinValue(1)
              .setMaxValue(15)
          )
      )
      .toJSON(),
  ];

  const rest = new REST({ version: '10' }).setToken(TOKEN);

  try {
    if (GUILD_ID) {
      await rest.put(Routes.applicationGuildCommands(CLIENT_ID, GUILD_ID), { body: commands });
      console.log('✅ Slash commands înregistrate pe server');
    } else {
      await rest.put(Routes.applicationCommands(CLIENT_ID), { body: commands });
      console.log('✅ Slash commands înregistrate global');
    }
  } catch (err) {
    console.error('Eroare la înregistrarea comenzilor:', err);
  }

  setInterval(checkFarmStops, CHECK_INTERVAL_MS);
});

client.on('messageCreate', async (message) => {
  if (message.author.bot && !message.webhookId) return;

  const data = extractFarmData(message);
  if (data) {
    channelFarmState.set(message.channel.id, {
      lastTimestamp: data.timestamp.getTime(),
      notified: false,
      channel: message.channel,
    });
  }

  // Prefix commands
  if (!message.content.toLowerCase().startsWith('!farm')) return;

  const args = message.content.slice(5).trim().split(/\s+/);
  const isHistory = args[0]?.toLowerCase() === 'history';

  await message.channel.sendTyping();

  if (isHistory) {
    const count = parseInt(args[1]) || 5;
    const farmMessages = await fetchFarmMessages(message.channel, 72);
    const sessions = groupIntoSessions(farmMessages).reverse().slice(0, count);

    if (sessions.length === 0) {
      return message.reply('❌ Nu am găsit nicio sesiune de farm.');
    }

    const embed = new EmbedBuilder()
      .setColor(0x5865f2)
      .setTitle(`📜 Istoric Farm (ultimele ${sessions.length} sesiuni)`)
      .setTimestamp();

    sessions.forEach((session, index) => {
      const start = session[0];
      const end = session[session.length - 1];
      const totalCoins = session.reduce((s, m) => s + m.coins, 0);
      const durationMin = Math.max(1, Math.round((end.timestamp - start.timestamp) / 60000));
      const h = Math.floor(durationMin / 60);
      const m = durationMin % 60;

      embed.addFields({
        name: `#${index + 1} — ${formatTime(start.timestamp)}`,
        value: `💰 **${totalCoins.toLocaleString('ro-RO')}** coins • ⏱️ ${h > 0 ? h + 'h ' : ''}${m} min • ${session.length} mesaje`,
        inline: false
      });
    });

    return message.reply({ embeds: [embed] });
  }

  // !farm [ore]
  const hours = args[0] ? parseInt(args[0], 10) : null;
  const farmMessages = await fetchFarmMessages(message.channel, hours);
  const embed = buildSummaryEmbed(farmMessages, hours);
  const buttons = createButtons(hours);

  await message.reply({ embeds: [embed], components: [buttons] });
});

client.on('interactionCreate', async (interaction) => {
  if (interaction.isChatInputCommand() && interaction.commandName === 'farm') {
    const sub = interaction.options.getSubcommand();

    if (sub === 'now') {
      const hours = interaction.options.getInteger('ore');
      await interaction.deferReply();

      const farmMessages = await fetchFarmMessages(interaction.channel, hours);
      const embed = buildSummaryEmbed(farmMessages, hours);
      const buttons = createButtons(hours);

      await interaction.editReply({ embeds: [embed], components: [buttons] });
    }

    if (sub === 'history') {
      const count = interaction.options.getInteger('numar') || 5;
      await interaction.deferReply();

      const farmMessages = await fetchFarmMessages(interaction.channel, 72);
      const sessions = groupIntoSessions(farmMessages).reverse().slice(0, count);

      if (sessions.length === 0) {
        return interaction.editReply({ content: '❌ Nu am găsit nicio sesiune de farm în ultimele zile.' });
      }

      const embed = new EmbedBuilder()
        .setColor(0x5865f2)
        .setTitle(`📜 Istoric Farm (ultimele ${sessions.length} sesiuni)`)
        .setTimestamp();

      sessions.forEach((session, index) => {
        const start = session[0];
        const end = session[session.length - 1];
        const totalCoins = session.reduce((s, m) => s + m.coins, 0);
        const durationMin = Math.max(1, Math.round((end.timestamp - start.timestamp) / 60000));
        const h = Math.floor(durationMin / 60);
        const m = durationMin % 60;

        embed.addFields({
          name: `#${index + 1} — ${formatTime(start.timestamp)}`,
          value: `💰 **${totalCoins.toLocaleString('ro-RO')}** coins • ⏱️ ${h > 0 ? h + 'h ' : ''}${m} min • ${session.length} mesaje`,
          inline: false
        });
      });

      await interaction.editReply({ embeds: [embed] });
    }
  }

  // Butoane
  if (interaction.isButton()) {
    const [action, type, hoursStr] = interaction.customId.split('_');
    const hours = hoursStr === 'all' ? null : parseInt(hoursStr, 10);

    if (action === 'recalc') {
      await interaction.deferUpdate();
      const farmMessages = await fetchFarmMessages(interaction.channel, hours);
      const embed = buildSummaryEmbed(farmMessages, hours);
      const buttons = createButtons(hours);
      await interaction.editReply({ embeds: [embed], components: [buttons] });
    }

    if (action === 'export') {
      await interaction.deferReply({ ephemeral: true });
      const farmMessages = await fetchFarmMessages(interaction.channel, hours);

      if (farmMessages.length === 0) {
        return interaction.editReply({ content: '❌ Nu există date de exportat.' });
      }

      let filePath;
      if (type === 'csv') {
        filePath = await generateCSV(farmMessages);
      } else {
        filePath = await generateExcel(farmMessages);
      }

      const attachment = new AttachmentBuilder(filePath);
      await interaction.editReply({
        content: `✅ Export ${type.toUpperCase()} gata! (${farmMessages.length} înregistrări)`,
        files: [attachment],
      });

      setTimeout(() => {
        try { fs.unlinkSync(filePath); } catch {}
      }, 30000);
    }
  }
});

async function checkFarmStops() {
  const now = Date.now();
  const timeoutMs = STOP_TIMEOUT_MINUTES * 60 * 1000;

  for (const [channelId, state] of channelFarmState.entries()) {
    if (state.notified) continue;

    if (now - state.lastTimestamp > timeoutMs) {
      try {
        const embed = new EmbedBuilder()
          .setColor(0xff9900)
          .setTitle('🛑 Farm oprit')
          .setDescription(
            `Nu a mai venit niciun mesaj cu coins de **${STOP_TIMEOUT_MINUTES} minute**.\nUltimul mesaj: <t:${Math.floor(state.lastTimestamp / 1000)}:R>`
          )
          .setTimestamp();

        await state.channel.send({ embeds: [embed] });
        state.notified = true;
      } catch (err) {
        console.error(`Nu am putut notifica pe canalul ${channelId}:`, err.message);
      }
    }
  }
}

// ========== START ==========
if (!TOKEN || !CLIENT_ID) {
  console.error('❌ Lipsește TOKEN sau CLIENT_ID în Environment Variables!');
  process.exit(1);
}

client.login(TOKEN);
