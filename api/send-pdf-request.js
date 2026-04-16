export default async function handler(req, res) {
  // --- CORS ---
  res.setHeader("Access-Control-Allow-Origin", "*")
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS")
  res.setHeader("Access-Control-Allow-Headers", "Content-Type")

  if (req.method === "OPTIONS") {
    return res.status(200).end()
  }

  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" })
  }

  try {
    const body = req.body || {}
    const type = body.type
    const articleId = body.articleId || "default-article"
    const contact = body.contact || {}

    const name = (contact.name || "").trim()
    const company = (contact.company || "").trim()
    const role = (contact.role || "").trim()
    const email = (contact.email || "").trim()

    if (type !== "pdf_request") {
      return res.status(400).json({ error: "Invalid request type" })
    }

    if (!email) {
      return res.status(400).json({ error: "Email is required" })
    }

    // --- ARTICLE CONFIG ---
    const ARTICLES = {
      "virtual-rtu-whitepaper": {
        subject: "Your Virtual RTU Whitepaper",
        pdfUrl:
          "https://www.resonant-grid.com/pdfs/virtual-rtu-whitepaper.pdf",
        filename: "virtual-rtu-whitepaper.pdf",
      },
      "default-article": {
        subject: "Your requested document",
        pdfUrl: "https://www.resonant-grid.com/pdfs/default.pdf",
        filename: "document.pdf",
      },
    }

    const article = ARTICLES[articleId] || ARTICLES["default-article"]

    // --- FETCH PDF ---
    const pdfResponse = await fetch(article.pdfUrl)

    if (!pdfResponse.ok) {
      throw new Error(`Failed to fetch PDF from ${article.pdfUrl}`)
    }

    const arrayBuffer = await pdfResponse.arrayBuffer()
    const pdfBase64 = Buffer.from(arrayBuffer).toString("base64")

    // --- SEND EMAIL ---
    await sendEmail({
      to: email,
      subject: article.subject,
      html: buildEmailHtml({ name, company, role }),
      attachments: [
        {
          filename: article.filename,
          content: pdfBase64,
          encoding: "base64",
          type: "application/pdf",
        },
      ],
    })

    // --- OPTIONAL: LOG LEAD ---
    console.log("PDF request received", {
      articleId,
      name,
      company,
      role,
      email,
      submittedAt: body.submittedAt,
    })

    return res.status(200).json({ ok: true })
  } catch (error) {
    console.error("send-pdf-request error:", error)

    return res.status(500).json({
      error: "Failed to send PDF",
      detail: error.message,
    })
  }
}

//
// --- EMAIL TEMPLATE ---
//
function buildEmailHtml({ name }) {
  return `
    <div style="font-family: Arial, sans-serif; line-height: 1.5;">
      <p>Hi${name ? " " + name : ""},</p>

      <p>
        Thanks for your interest. Your requested document is attached.
      </p>

      <p>
        If you'd like to discuss this further, feel free to reply to this email.
      </p>

      <p>
        Regards,<br/>
        Resonant Grid
      </p>
    </div>
  `
}

//
// --- EMAIL SENDER ---
// IMPORTANT: Replace this with your existing ROI webhook email logic
//
async function sendEmail({ to, subject, html, attachments }) {
  /**
   * ============================================
   * OPTION A — RESEND (recommended)
   * ============================================
   */

  if (process.env.RESEND_API_KEY) {
    const response = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${process.env.RESEND_API_KEY}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        from: "Resonant Grid <info@resonant-grid.com>",
        to: [to],
        subject,
        html,
        attachments,
      }),
    })

    if (!response.ok) {
      const text = await response.text()
      throw new Error(`Resend error: ${text}`)
    }

    return
  }

  /**
   * ============================================
   * OPTION B — FAIL FAST (if not configured)
   * ============================================
   */

  throw new Error(
    "No email provider configured. Add RESEND_API_KEY or reuse ROI email logic."
  )
}
