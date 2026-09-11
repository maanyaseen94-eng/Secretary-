// ملف: api/draft-reply.js
// هذا المسار يصبح تلقائياً رابط: https://اسم-مشروعك.vercel.app/api/draft-reply

const Anthropic = require("@anthropic-ai/sdk");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
} = require("docx");

module.exports = async (req, res) => {
  // السماح لموقع السكرتارية (GitHub Pages) يتصل بهذا الرابط
  // بدّل * لاحقاً برابط موقعك تحديداً لأمان أفضل، مثلاً:
  // res.setHeader("Access-Control-Allow-Origin", "https://maanyaseen94-eng.github.io");
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  // المتصفح يرسل طلب OPTIONS أول تلقائياً (preflight) — لازم نرد عليه فاضي
  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  if (req.method !== "POST") {
    return res.status(405).json({ error: "الطريقة غير مسموحة" });
  }

  const { requestSummary, requestType, entityName, entityTitle } = req.body || {};

  if (!requestSummary || requestSummary.trim().length < 3) {
    return res.status(400).json({ error: "لازم تكتب ملخص الطلب أول" });
  }

  const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

  const prompt = `
انت مساعد إداري بمكتب نائب برلماني عراقي. اكتب مسودة رد رسمي بالعربية الفصحى
على المراسلة التالية، بأسلوب إداري رسمي مناسب للمكاتبات الحكومية العراقية.

الجهة المرسل إليها (اسم/صفة): ${entityName || "غير محدد"}
لقب أو منصب الجهة (وزير، مدير عام، دائرة، مواطن...): ${entityTitle || "غير محدد"}
نوع المراسلة: ${requestType || "مراسلة عامة"}
ملخص الطلب: ${requestSummary}

اختر صيغة المخاطبة الرسمية المناسبة تماماً حسب لقب ومنصب الجهة أعلاه
(مثلاً: السيد الوزير المحترم، دائرة ... المحترمة، السيد المواطن المحترم، إلخ).
اكتب الرد فقط، بدون أي شرح إضافي، وابدأ مباشرة بالمخاطبة الرسمية.
`;

  try {
    const response = await client.messages.create({
      model: "claude-sonnet-4-5",
      max_tokens: 1024,
      messages: [{ role: "user", content: prompt }],
    });

    const draftText = response.content[0].text;
    const docxBase64 = await buildWordDocument(draftText);

    return res.status(200).json({ draft: draftText, docxBase64 });
  } catch (error) {
    console.error("خطأ بالاتصال مع Claude API:", error);
    return res.status(500).json({ error: "صار خطأ أثناء توليد الرد، حاول مرة ثانية" });
  }
};

async function buildWordDocument(draftText) {
  const lines = draftText.split("\n").filter((line) => line.trim().length > 0);

  const paragraphs = lines.map(
    (line) =>
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        bidirectional: true,
        children: [
          new TextRun({ text: line, font: "Arial", size: 24 }),
        ],
        spacing: { after: 200 },
      })
  );

  const doc = new Document({ sections: [{ children: paragraphs }] });
  const buffer = await Packer.toBuffer(doc);
  return buffer.toString("base64");
}
