const axios = require("axios");
const ExcelJS = require("exceljs");
const sharp = require("sharp");
const { Telegraf } = require("telegraf");
const { HttpsProxyAgent } = require("https-proxy-agent");

// ===================== SETTINGS ======================

const BOT_TOKEN = process.env.BOT_TOKEN || "PASTE_YOUR_BOT_TOKEN_HERE";

if (!BOT_TOKEN || BOT_TOKEN === "PASTE_YOUR_BOT_TOKEN_HERE") {
  throw new Error("BOT_TOKEN не задан. Укажи BOT_TOKEN в env.");
}

const TG_PROXY = (process.env.TG_PROXY || "").trim() || null;

const IMAGE_BASES = [
  "https://goodsmanager.easykassa.ru/images",
  "https://storage.easykassa.ru/images",
];

const SHOP_CONFIGS = {
  shop1: {
    name: "Магазин 1 (pickme)",
    origin: "https://pickme.evosell.ru",
    catalogUrl: "https://pickme.evosell.ru/products",
    userId: "01-000000011435223",
    cookie:
      "_ga=GA1.2.467391203.1757935954; _ga_RWP4RVD64M=GS2.2.s1766256569$o104$g0$t1766256569$j60$l0$h0; data-domain=ru; appLanguage=ru-RU",
    referer: "https://pickme.evosell.ru/catalog?c=",
  },
  shop2: {
    name: "Магазин 2 (optskladpickme)",
    origin: "https://optskladpickme.evosell.ru",
    catalogUrl: "https://optskladpickme.evosell.ru/products",
    userId: "01-000000012259862",
    cookie:
      "_ga=GA1.2.304634254.1762007835; data-domain=ru; _gid=GA1.2.1956320999.1766516893; _ga_RWP4RVD64M=GS2.2.s1766579288$o40$g0$t1766579288$j60$l0$h0; appLanguage=ru-RU",
    referer: "https://optskladpickme.evosell.ru/catalog?c=",
  },
};

const RESET_MODE_AFTER_PROCESS = true;

// режимы пользователей
const userModes = new Map();

// ===================== HTTP ======================

const axiosInstance = axios.create({
  timeout: 35000,
  maxRedirects: 5,
  validateStatus: () => true,
  ...(TG_PROXY ? { httpsAgent: new HttpsProxyAgent(TG_PROXY) } : {}),
});

// ===================== HELPERS ======================

function extractProductsFromResponse(data) {
  if (Array.isArray(data)) {
    const out = [];
    for (const x of data) {
      if (x && typeof x === "object" && x._source && typeof x._source === "object") {
        out.push(x._source);
      } else if (x && typeof x === "object") {
        out.push(x);
      }
    }
    return out;
  }

  if (!data || typeof data !== "object") {
    return [];
  }

  if (data.hits && typeof data.hits === "object") {
    const hits = data.hits.hits || [];
    const out = [];
    for (const h of hits) {
      if (h && typeof h === "object") {
        const src = h._source;
        out.push(src && typeof src === "object" ? src : h);
      }
    }
    return out;
  }

  for (const key of ["items", "data", "results", "products", "content"]) {
    if (Array.isArray(data[key])) {
      return data[key].map((x) => (x && typeof x === "object" ? x._source || x : x));
    }
  }

  return [];
}

function normalizeArticle(article) {
  let a = String(article || "").trim();
  if (a.endsWith(".0")) {
    a = a.slice(0, -2);
  }
  return a;
}

async function fetchProductByArticle(shopCfg, article) {
  article = normalizeArticle(article);
  if (!article) return null;

  const headers = {
    accept: "application/json",
    "content-type": "application/json",
    origin: shopCfg.origin,
    referer: shopCfg.referer,
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    cookie: shopCfg.cookie,
  };

  const payload = {
    query: {
      bool: {
        must: [
          { term: { "user_id.keyword": shopCfg.userId } },
          {
            dis_max: {
              queries: [
                { match_phrase: { name: { query: article, boost: 5 } } },
                { match_phrase: { description: { query: article, boost: 4 } } },
                { match_phrase: { article_number: { query: article, boost: 5 } } },
                { match_phrase: { code: { query: article, boost: 5 } } },
                { match_phrase_prefix: { name: { query: article, boost: 2 } } },
                { match_phrase_prefix: { description: { query: article, boost: 2 } } },
                { match_phrase_prefix: { article_number: { query: article, boost: 2 } } },
                { match_phrase_prefix: { code: { query: article, boost: 2 } } },
                {
                  multi_match: {
                    query: article,
                    fields: ["name^3", "description", "article_number", "code"],
                    type: "best_fields",
                    operator: "and",
                    boost: 2,
                  },
                },
              ],
            },
          },
        ],
      },
    },
    size: 10,
  };

  try {
    const resp = await axiosInstance.post(shopCfg.catalogUrl, payload, {
      headers,
      timeout: 12000,
    });

    if (resp.status !== 200) {
      console.warn(`⚠️ ${shopCfg.catalogUrl} вернул ${resp.status} по артикулу ${article}`);
      return null;
    }

    const data = resp.data;
    if (!data || typeof data !== "object") {
      console.warn(`⚠️ API вернул не JSON по ${article}`);
      return null;
    }

    const products = extractProductsFromResponse(data);
    if (!products.length) return null;

    for (const p of products) {
      const code = String(p?.code || "").trim();
      const artn = String(p?.article_number || "").trim();
      if (code === article || artn === article) {
        return p;
      }
    }

    return products[0];
  } catch (e) {
    console.warn(`⚠️ Ошибка запроса товара по артикулу ${article}: ${e.message}`);
    return null;
  }
}

function buildCandidateImageUrls(fileEntry) {
  const urls = [];
  if (!fileEntry || typeof fileEntry !== "object") return urls;

  for (const key of ["uriMid", "uriMin", "url", "file", "src", "path", "original"]) {
    const raw = fileEntry[key];
    if (!raw) continue;

    let val = String(raw).trim();

    if (val.startsWith("http://") || val.startsWith("https://")) {
      urls.push(val);
      continue;
    }

    val = val.replace(/^\/+/, "");
    for (const base of IMAGE_BASES) {
      urls.push(`${base}/${val}`);
    }
  }

  return [...new Set(urls)];
}

async function downloadImage(shopCfg, url) {
  try {
    const resp = await axiosInstance.get(url, {
      responseType: "arraybuffer",
      timeout: 12000,
      headers: {
        "user-agent": "Mozilla/5.0",
        accept: "image/avif,image/webp,image/apng,image/*,*/*;q=0.8",
        referer: shopCfg.referer,
      },
    });

    if (resp.status !== 200) {
      return null;
    }

    const ctype = String(resp.headers["content-type"] || "").toLowerCase();
    if (ctype.includes("text/html")) {
      return null;
    }

    const content = Buffer.from(resp.data);
    if (!content || content.length < 50) {
      return null;
    }

    return content;
  } catch (e) {
    console.warn(`⚠️ Ошибка GET ${url}: ${e.message}`);
    return null;
  }
}

async function makePngThumbnail(imgBuffer, maxW = 220, maxH = 220) {
  try {
    return await sharp(imgBuffer)
      .rotate()
      .resize({
        width: maxW,
        height: maxH,
        fit: "inside",
        withoutEnlargement: true,
      })
      .png()
      .toBuffer();
  } catch (e) {
    console.warn(`⚠️ Ошибка обработки картинки: ${e.message}`);
    return null;
  }
}

async function getImageForArticle(shopCfg, article) {
  const product = await fetchProductByArticle(shopCfg, article);
  if (!product) {
    return { imageBuffer: null, triedUrl: null };
  }

  const images = product.images || {};
  const files = images.files || [];
  if (!files.length) {
    return { imageBuffer: null, triedUrl: null };
  }

  let first = files[0];
  if (typeof first === "string") {
    first = { uriMid: first };
  }

  const candidates = buildCandidateImageUrls(first);
  if (!candidates.length) {
    return { imageBuffer: null, triedUrl: null };
  }

  for (const url of candidates) {
    const imgBytes = await downloadImage(shopCfg, url);
    if (!imgBytes) continue;

    const pngBuffer = await makePngThumbnail(imgBytes);
    if (pngBuffer) {
      return { imageBuffer: pngBuffer, triedUrl: url };
    }
  }

  return { imageBuffer: null, triedUrl: candidates[0] };
}

function findHeaderColumnIndex(worksheet, headerName) {
  const row1 = worksheet.getRow(1);
  const count = Math.max(row1.cellCount, worksheet.columnCount);

  for (let i = 1; i <= count; i++) {
    const val = row1.getCell(i).value;
    if (val && String(val).trim().toLowerCase() === headerName.toLowerCase()) {
      return i;
    }
  }

  return null;
}

function cellValueToString(value) {
  if (value == null) return "";

  if (typeof value === "object") {
    if (value.text) return String(value.text);
    if (Array.isArray(value.richText)) {
      return value.richText.map((x) => x.text || "").join("");
    }
    if (value.result != null) return String(value.result);
  }

  return String(value);
}

async function processExcelBuffer(excelBuffer, shopKey) {
  const shopCfg = SHOP_CONFIGS[shopKey];
  const workbook = new ExcelJS.Workbook();

  await workbook.xlsx.load(excelBuffer);
  const worksheet = workbook.worksheets[0];

  if (!worksheet) {
    throw new Error("Не удалось открыть лист Excel");
  }

  const articleCol = findHeaderColumnIndex(worksheet, "Артикул");
  if (!articleCol) {
    throw new Error("В первой строке нет столбца 'Артикул'");
  }

  const photoCol = worksheet.columnCount + 1;
  worksheet.getCell(1, photoCol).value = "Фото";
  worksheet.getColumn(photoCol).width = 35;

  for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber++) {
    try {
      const artRaw = worksheet.getCell(rowNumber, articleCol).value;
      const artVal = normalizeArticle(cellValueToString(artRaw));

      if (!artVal) continue;

      console.log(`Обработка строки ${rowNumber}, артикул: ${artVal}`);

      const { imageBuffer, triedUrl } = await getImageForArticle(shopCfg, artVal);

      if (imageBuffer) {
        const imageId = workbook.addImage({
          buffer: imageBuffer,
          extension: "png",
        });

        worksheet.getRow(rowNumber).height = 95;

        worksheet.addImage(imageId, {
          tl: { col: photoCol - 1 + 0.1, row: rowNumber - 1 + 0.1 },
          ext: { width: 160, height: 120 },
          editAs: "oneCell",
        });
      } else {
        worksheet.getCell(rowNumber, photoCol).value = triedUrl || `NO_IMAGE:${artVal}`;
      }
    } catch (rowError) {
      console.warn(`⚠️ Ошибка на строке ${rowNumber}: ${rowError.message}`);
      worksheet.getCell(rowNumber, photoCol).value = `ERROR:${rowError.message}`;
    }
  }

  const out = await workbook.xlsx.writeBuffer();
  return Buffer.from(out);
}

// ===================== TELEGRAM ======================

const botOptions = TG_PROXY
  ? {
      telegram: {
        agent: new HttpsProxyAgent(TG_PROXY),
      },
    }
  : {};

const bot = new Telegraf(BOT_TOKEN, {
  ...botOptions,
  handlerTimeout: 600000,
});

bot.catch((err, ctx) => {
  console.error("Unhandled error while processing", ctx.update);
  console.error(err);
});

bot.start(async () => {
  // игнорируем /start как в оригинале
});

bot.command("command1", async (ctx) => {
  userModes.set(ctx.from.id, "shop1");
  await ctx.replyWithMarkdown(
    `Ок ✅ Режим включён: *${SHOP_CONFIGS.shop1.name}*\nПришли Excel (.xlsx).`
  );
});

bot.command("command2", async (ctx) => {
  userModes.set(ctx.from.id, "shop2");
  await ctx.replyWithMarkdown(
    `Ок ✅ Режим включён: *${SHOP_CONFIGS.shop2.name}*\nПришли Excel (.xlsx).`
  );
});

bot.on("document", async (ctx) => {
  const doc = ctx.message.document;
  const fileName = doc.file_name || "";

  if (!(fileName.endsWith(".xlsx") || fileName.endsWith(".xlsm"))) {
    await ctx.reply("Пришли .xlsx 🙏");
    return;
  }

  const shopKey = userModes.get(ctx.from.id);
  if (!SHOP_CONFIGS[shopKey]) {
    await ctx.reply("Сначала выбери режим: /command1 или /command2");
    return;
  }

  const shopName = SHOP_CONFIGS[shopKey].name;
  await ctx.replyWithMarkdown(`Обрабатываю файл… 🛠\nИсточник: *${shopName}*`);
  console.log(`User ${ctx.from.id} processing ${fileName} using ${shopKey}`);

  try {
    const fileLink = await ctx.telegram.getFileLink(doc.file_id);
    const fileResp = await axiosInstance.get(fileLink.href, {
      responseType: "arraybuffer",
      timeout: 30000,
    });

    const excelBuffer = Buffer.from(fileResp.data);
    const resultBuffer = await processExcelBuffer(excelBuffer, shopKey);

    const outName = `result_${shopKey}_${fileName.replace(/\.xlsm$/i, ".xlsx")}`;

    await ctx.replyWithDocument(
      {
        source: resultBuffer,
        filename: outName,
      },
      {
        caption: `Готово ✅ (${shopName})`,
      }
    );
  } catch (e) {
    console.error("Ошибка при обработке файла:", e);
    await ctx.reply(`Ошибка: ${e.message || e}`);
  } finally {
    if (RESET_MODE_AFTER_PROCESS) {
      userModes.delete(ctx.from.id);
    }
  }
});

// ===================== START ======================

async function main() {
  console.log("Бот запущен");
  if (TG_PROXY) {
    console.log(`Использую TG_PROXY=${TG_PROXY}`);
  }

  await bot.launch();

  process.once("SIGINT", () => bot.stop("SIGINT"));
  process.once("SIGTERM", () => bot.stop("SIGTERM"));
}

main().catch((err) => {
  console.error("Фатальная ошибка:", err);
  process.exit(1);
});