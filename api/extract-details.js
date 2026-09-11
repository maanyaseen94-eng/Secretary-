// ملف: api/extract-details.js
// يستقبل صورة أو PDF للمراسلة، ويستخرج منه: الجهة، الموضوع، نوع المراسلة

const Anthropic = require("@anthropic-ai/sdk");

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

  const { fileBase64, mediaType } = req.body || {};

  if (!fileBase64 || !mediaType) {
    return res.status(400).json({ error: "لازم ترفع ملف أول" });
  }

  const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

  // نحدد نوع المحتوى المرسل لـ Claude حسب نوع الملف (صورة أو PDF)
  const isPdf = mediaType === "application/pdf";
  const fileContentBlock = isPdf
    ? { type: "document", source: { type: "base64", media_type: mediaType, data: fileBase64 } }
    : { type: "image", source: { type: "base64", media_type: mediaType, data: fileBase64 } };

  const instructionText = `
هذي صورة أو مستند لمراسلة رسمية وصلت لمكتب نائب برلماني عراقي.
اقرأ المحتوى واستخرج المعلومات التالية بدقة، وأرجعها بصيغة JSON فقط بدون أي نص إضافي:

{
  "entityName": "اسم الجهة المرسلة (مثال: وزارة الصحة، أو اسم شخص إذا كانت مراسلة من مواطن)",
  "entityTitle": "لقب أو منصب الجهة (مثال: السيد الوزير، دائرة، مواطن)",
  "requestType": "اختر واحد بالضبط من هذي الخيارات: طلب موافقة | استفسار | شكوى | دعوة | (اتركه فارغ لو ما ينطبق)",
  "requestSummary": "ملخص موجز وواضح لمضمون المراسلة والطلب المذكور فيها، بجملتين لثلاث جمل"
}

إذا ما قدرت تقرأ معلومة معينة بوضوح، اكتب لها قيمة فارغة "". لا تضيف أي شرح أو نص خارج كائن JSON.
`;

  try {
    const response = await client.messages.create({
      model: "claude-sonnet-4-5",
      max_tokens: 1024,
      messages: [
        {
          role: "user",
          content: [fileContentBlock, { type: "text", text: instructionText }],
        },
      ],
    });

    const rawText = response.content[0].text.trim();
    // نتأكد نشيل أي علامات ```json``` لو Claude ضافها بالغلط
    const cleanText = rawText.replace(/```json|```/g, "").trim();
    const extracted = JSON.parse(cleanText);

    return res.status(200).json(extracted);
  } catch (error) {
    console.error("خطأ أثناء استخراج التفاصيل:", error);
    return res.status(500).json({ error: "صار خطأ أثناء قراءة الملف، حاول مرة ثانية" });
  }
};
