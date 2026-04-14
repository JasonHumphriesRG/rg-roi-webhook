import { Resend } from "resend"
import { google } from "googleapis"
import PDFDocument from "pdfkit"
import { PassThrough, Readable } from "stream"

let resend = null

function setCors(res) {
  res.setHeader("Access-Control-Allow-Origin", "*")
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS")
  res.setHeader("Access-Control-Allow-Headers", "Content-Type")
}

function isUsMarket(market) {
  const m = String(market || "").toLowerCase()
  return m.includes("us")
}

function getPageSize(market) {
  return isUsMarket(market) ? "LETTER" : "A4"
}

function formatCurrency(value, currency = "GBP") {
  return new Intl.NumberFormat("en-GB", {
    style: "currency",
    currency,
    maximumFractionDigits: 0,
  }).format(value || 0)
}

function formatMultiple(value) {
  return `${Number(value || 0).toFixed(1)}x`
}

function formatMonths(value) {
  if (!isFinite(value)) return "—"
  return `${value.toFixed(1)} months`
}

async function buildPdfBuffer(data) {
  return new Promise((resolve, reject) => {
    try {
      const doc = new PDFDocument({
        size: getPageSize(data.calculator.market),
        margin: 30,
      })

      const stream = new PassThrough()
      const chunks = []

      stream.on("data", (c) => chunks.push(c))
      stream.on("end", () => resolve(Buffer.concat(chunks)))
      doc.pipe(stream)

      const currency = data.calculator.currency
      const inputs = data.calculator.inputs
      const a = data.calculator.assumptions

      // === V3 CALCULATOR (aligned to spreadsheet) ===
      const reliability =
        inputs.annualRegulatedRevenue *
        (a.reliabilityExposurePct / 100) *
        (a.underClaimPct / 100)

      const misattribution =
        inputs.numberOfMeters *
        (inputs.eventsPerMeter || 0) *
        (a.misattributionPct / 100) *
        inputs.avgCostPerMisclassifiedEvent

      const telecom =
        inputs.numberOfMeters *
        inputs.monthlyCommsCostPerMeter *
        12 *
        (a.telecomRecoveryPct / 100)

      const operational =
        inputs.numberOfMeters *
        (inputs.reportingCostPerMeter || 0) *
        (a.operationalSavingsPct / 100)

      const total =
        reliability + misattribution + telecom + operational

      const fee =
        inputs.numberOfMeters * inputs.reePricePerMeterPerYear

      const net = total - fee
      const roi = total / fee
      const payback = fee / (total / 12)

      // === HEADER ===
      doc.fontSize(16).text("REE ROI Summary", { align: "left" })
      doc.moveDown()

      doc.fontSize(10).text(`Utility: ${data.calculator.utilityName}`)
      doc.text(`Market: ${data.calculator.market}`)
      doc.text(`Submitted: ${new Date().toLocaleString()}`)
      doc.moveDown()

      // === KPI ===
      doc.fontSize(12).text("Summary")
      doc.text(`Total Value: ${formatCurrency(total, currency)}`)
      doc.text(`REE Fee: ${formatCurrency(fee, currency)}`)
      doc.text(`Net Benefit: ${formatCurrency(net, currency)}`)
      doc.text(`ROI: ${formatMultiple(roi)}`)
      doc.text(`Payback: ${formatMonths(payback)}`)
      doc.moveDown()

      // === VALUE BUILD ===
      doc.fontSize(12).text("Value Build-Up")
      doc.text(`Reliability: ${formatCurrency(reliability, currency)}`)
      doc.text(`Misattribution: ${formatCurrency(misattribution, currency)}`)
      doc.text(`Telecom: ${formatCurrency(telecom, currency)}`)
      doc.text(`Operational: ${formatCurrency(operational, currency)}`)
      doc.moveDown()

      // === ASSUMPTIONS ===
      doc.fontSize(12).text("Assumptions")
      Object.entries(a).forEach(([k, v]) => {
        doc.text(`${k}: ${v}%`)
      })

      doc.end()
    } catch (err) {
      reject(err)
    }
  })
}

export default async function handler(req, res) {
  setCors(res)

  if (req.method === "OPTIONS") return res.status(200).end()

  try {
    if (!resend) resend = new Resend(process.env.RESEND_API_KEY)

    const data = req.body
    const email = data?.contact?.email

    if (!email) {
      return res.status(400).json({ error: "Missing email" })
    }

    const pdf = await buildPdfBuffer(data)

    // send email
    await resend.emails.send({
      from: "sales@resonant-grid.com",
      to: email,
      subject: "ROI Summary",
      html: `<p>Thanks for your submission. Your ROI summary is attached.</p>`,
      attachments: [
        {
          filename: "roi-summary.pdf",
          content: pdf,
        },
      ],
    })

    return res.status(200).json({ success: true })
  } catch (err) {
    console.error(err)
    return res.status(500).json({ error: err.message })
  }
}
