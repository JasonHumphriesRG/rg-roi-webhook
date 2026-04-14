import { Resend } from "resend"
import PDFDocument from "pdfkit"
import { PassThrough } from "stream"

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

      const inputs = data.calculator.inputs
      const a = data.calculator.assumptions
      const currency = data.calculator.currency

      // === CALCS ===
      function calcROI(s) {
        const reliability =
          inputs.annualRegulatedRevenue *
          (s.reliabilityExposurePct / 100) *
          (s.underClaimPct / 100)

        const misattribution =
          inputs.annualOutageEvents *
          (s.misattributionPct / 100) *
          inputs.avgCostPerMisclassifiedEvent

        const telecom =
          inputs.numberOfMeters *
          inputs.monthlyCommsCostPerMeter *
          12 *
          (s.telecomRecoveryPct / 100)

        const operational =
          (inputs.annualInternalFTECost +
            inputs.annualConsultantCost +
            inputs.annualManagementOverhead +
            inputs.annualAuditRemediationCost) *
          (s.operationalSavingsPct / 100)

        const total =
          reliability + misattribution + telecom + operational

        const fee =
          inputs.numberOfMeters * inputs.reePricePerMeterPerYear

        return fee > 0 ? total / fee : 0
      }

      const base = calcROI(a)
      const low = calcROI(data.calculator.scenarioLow || a)
      const high = calcROI(data.calculator.scenarioHigh || a)

      // === HEADER ===
      doc.fontSize(16).text("REE ROI Summary")
      doc.fontSize(8).text("Model Version: V3", { align: "right" })
      doc.moveDown()

      // === KPI ===
      doc.fontSize(12).text("Summary")
      doc.text(`ROI: ${formatMultiple(base)}`)
      doc.text(`Range: ${formatMultiple(low)} – ${formatMultiple(high)}`)
      doc.moveDown()

      // === BUSINESS INPUTS ===
      doc.fontSize(12).text("Business Inputs (User Supplied)")
      doc.text(`Utility: ${data.calculator.utilityName}`)
      doc.text(`Market: ${data.calculator.market}`)
      doc.text(`Meters: ${inputs.numberOfMeters}`)
      doc.text(`Revenue: ${formatCurrency(inputs.annualRegulatedRevenue, currency)}`)
      doc.text(`Events: ${inputs.annualOutageEvents}`)
      doc.moveDown()

      // === ASSUMPTIONS ===
      doc.fontSize(12).text("Model Assumptions")
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

    await resend.emails.send({
      from: "sales@resonant-grid.com",
      to: email,
      subject: "ROI Summary",
      html: `<p>Your ROI summary is attached.</p>`,
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
