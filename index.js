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
const STOP_TIMEOUT_MINUTES = 10;
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

// ========== KEEP-ALIVE SERVER (pentru Render) ==========
const app = express();
const PORT = process.env.PORT || 3000;

app.get('/', (req, res) => {
  res.send('Farm Bot is online ✅');
});

app.listen(PORT, () => {
  console.log(`Keep-alive server running on port ${PORT}`);
});

// ========== FUNCȚII ==========

function extractFarmData(msg) {
  let text = msg.content || '';
  if (msg.embeds?.length > 0) {
    text = msg.embeds
      .map((e) => {
        return [
          e.title || '',
          e.description || '',
          ...(e.fields?.map((f) => `${f.name}: ${f.value}`) || []),
        ].join('\n');
      })
      .join('\n');
  }

  const coinsMatch = text.match(/Coins Collected in 3 Minutes[:\s]*(\d+)/i);
  if (!coinsMatch) return null;

  const estimatedMatch = text.match(/Estimated Coins per Hour[:\s]*(\d+)/i);
  const bagMatch = text.match(/Bag[:\s]*(.+)/i);

  return {
    coins: parseInt(coinsMatch[1], 10),
    estimated: estimatedMatch ? parseInt(estimatedMatch[1], 10) : null,
    bag: bagMatch ? bagMatch[1].trim() : 'Unknown',
    timestamp: msg.createdAt,
    id: msg.id,
    url: msg.url,
  };
}

async function fetchFarmMessages(channel, hours = null) {
  const limit = 100;
  let allMessages = [];
  let lastId = null;
  const maxFetch = 1000;

  while (allMessages.length < maxFetch) {
    const options = { limit };
    if (lastId) options.before = lastId;

    const messages = await channel.messages.fetch(options);
    if (messages.size === 0) break;

    allMessages.push(...messages.values());
    lastId = messages.last().id;

    if (hours) {
      const cutoff = Date.now() - hours * 60 * 60 * 1000;
      if (messages.last().createdTimestamp < cutoff) break;
    }
  }

  let farmMessages = allMessages
    .map(extractFarmData)
    .filter((d) => d !== null);

  if (hours) {
    const cutoff = Date.now() - hours * 60 * 60 * 1000;
    farmMessages = farmMessages.filter((m) => m.timestamp.getTime() >= cutoff);
  }

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

function buildSummaryEmbed(farmMessages, hours = null) {
  if (farmMessages.length === 0) {
    return new EmbedBuilder()
      .setColor(0xff0000)
      .setTitle('❌ Niciun farm găsit')
      .setDescription(
        hours
          ? `Nu am găsit mesaje de farm în ultimele **${hours} ore**.`
          : 'Nu am găsit mesaje de farm în istoricul recent.'
      );
  }

  const start = farmMessages[0];
  const end = farmMessages[farmMessages.length - 1];
  const totalCoins = farmMessages.reduce((sum, m) => sum + m.coins, 0);
  const durationMs = end.timestamp - start.timestamp;
  const durationMin = Math.round(durationMs / 60000);
  const hoursPart = Math.floor(durationMin / 60);
  const minutesPart = durationMin % 60;
  const avgPerHour = durationMin > 0 ? Math.round((totalCoins / durationMin) * 60) : 0;

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
  const header = 'Timestamp,Coins,Estimated per Hour,Bag,Message URL\n';
  const rows = farmMessages
    .map(m => `"${formatTime(m.timestamp)}",${m.coins},${m.estimated || ''},"${m.bag}","${m.url}"`)
    .join('\n');
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
    { header: 'Bag', key: 'bag', width: 15 },
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
  sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };

  const filePath = path.join(__dirname, `farm_${Date.now()}.xlsx`);
  await workbook.xlsx.writeFile(filePath);
  return filePath;
}

// ========== EVENTS ==========

client.once('ready', async () => {
  console.log(`✅ Bot online ca ${client.user.tag}`);

  const commands = [
    new SlashCommandBuilder()
      .setName('farm')
      .setDescription('Analizează mesajele de farm din canal')
      .addIntegerOption(opt =>
        opt.setName('ore')
          .setDescription('Ultimele X ore (lasă gol pentru tot istoricul recent)')
          .setMinValue(1)
          .setMaxValue(168)
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

  if (message.content.toLowerCase().startsWith('!farm')) {
    const args = message.content.slice(5).trim();
    const hours = args ? parseInt(args, 10) : null;

    await message.channel.sendTyping();
    const farmMessages = await fetchFarmMessages(message.channel, hours);
    const embed = buildSummaryEmbed(farmMessages, hours);
    const buttons = createButtons(hours);

    await message.reply({ embeds: [embed], components: [buttons] });
  }
});

client.on('interactionCreate', async (interaction) => {
  if (interaction.isChatInputCommand() && interaction.commandName === 'farm') {
    const hours = interaction.options.getInteger('ore');
    await interaction.deferReply();

    const farmMessages = await fetchFarmMessages(interaction.channel, hours);
    const embed = buildSummaryEmbed(farmMessages, hours);
    const buttons = createButtons(hours);

    await interaction.editReply({ embeds: [embed], components: [buttons] });
  }

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
            `Nu a mai venit niciun mesaj de farm de **${STOP_TIMEOUT_MINUTES} minute**.\nUltimul mesaj: <t:${Math.floor(state.lastTimestamp / 1000)}:R>`
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
