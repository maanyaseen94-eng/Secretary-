// ملف: api/draft-reply.js
// هذا المسار يصبح تلقائياً رابط: https://اسم-مشروعك.vercel.app/api/draft-reply
 
const Anthropic = require("@anthropic-ai/sdk");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
} = require("docx");
 
module.exports = async (req, res) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
 
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
 
  // نطلب من Claude يرجع الأجزاء منفصلة (JSON) بدل نص حر
  // عشان نتحكم إحنا بالتنسيق والشكل النهائي للكتاب، مو Claude
  const prompt = `
انت مساعد إداري بمكتب نائب برلماني عراقي. جهّز محتوى كتاب رسمي بالعربية الفصحى
رداً على المراسلة التالية، بأسلوب إداري رسمي مناسب للمكاتبات الحكومية العراقية.
 
الجهة المرسل إليها (اسم/صفة): ${entityName || "غير محدد"}
لقب أو منصب الجهة (وزير، مدير عام، دائرة، مواطن...): ${entityTitle || "غير محدد"}
نوع المراسلة: ${requestType || "مراسلة عامة"}
ملخص الطلب: ${requestSummary}
 
أرجع الرد بصيغة JSON فقط بدون أي نص خارج الكائن، بهذا الشكل بالضبط:
 
{
  "subject": "عنوان موجز لموضوع الكتاب (3-6 كلمات، بدون كلمة الموضوع نفسها)",
  "greeting": "صيغة المخاطبة الرسمية الكاملة حسب لقب ومنصب الجهة (مثال: السيد الوزير المحترم)",
  "bodyParagraphs": ["الفقرة الأولى (إشارة لموضوع المراسلة الواردة)", "الفقرة الثانية (الرد أو القرار أو الإجراء بالتفصيل)"],
  "closing": "صيغة ختامية رسمية مناسبة (مثال: مع التقدير، أو وتفضلوا بقبول فائق الاحترام)"
}
 
لا تضيف أي شرح، مقدمة، أو نص خارج كائن JSON.
`;
 
  try {
    const response = await client.messages.create({
      model: "claude-sonnet-4-5",
      max_tokens: 1500,
      messages: [{ role: "user", content: prompt }],
    });
 
    const rawText = response.content[0].text.trim();
    const cleanText = rawText.replace(/```json|```/g, "").trim();
    const parts = JSON.parse(cleanText);
 
    const draftText = buildPlainTextDraft(parts);
    const docxBase64 = await buildWordDocument(parts);
 
    return res.status(200).json({ draft: draftText, docxBase64 });
  } catch (error) {
    console.error("خطأ بالاتصال مع Claude API:", error);
    return res.status(500).json({ error: "صار خطأ أثناء توليد الرد، حاول مرة ثانية" });
  }
};
 
// نص مبسط يظهر بمربع المعاينة بالتطبيق
function buildPlainTextDraft(parts) {
  return [
    "العدد: ....................",
    "التاريخ: ....................",
    "",
    `م/ ${parts.subject || ""}`,
    "",
    parts.greeting || "",
    "",
    ...(parts.bodyParagraphs || []),
    "",
    parts.closing || "",
  ].join("\n");
}
 
// ملف Word فعلي بتنسيق رسمي: عناصر مهمة بخط عريض، فقرات المتن عادية
async function buildWordDocument(parts) {
  const boldLine = (text) =>
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      bidirectional: true,
      spacing: { after: 200 },
      children: [new TextRun({ text, font: "Arial", size: 24, bold: true })],
    });
 
  const normalLine = (text) =>
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      bidirectional: true,
      spacing: { after: 200 },
      children: [new TextRun({ text, font: "Arial", size: 24 })],
    });
 
  const paragraphs = [
    normalLine("العدد: ...................."),
    normalLine("التاريخ: ...................."),
    normalLine(""),
    boldLine(`م/ ${parts.subject || ""}`),
    normalLine(""),
    boldLine(parts.greeting || ""),
    normalLine(""),
    ...(parts.bodyParagraphs || []).map((p) => normalLine(p)),
    normalLine(""),
    normalLine(parts.closing || ""),
  ];
 
  const doc = new Document({ sections: [{ children: paragraphs }] });
  const buffer = await Packer.toBuffer(doc);
  return buffer.toString("base64");
}
 
