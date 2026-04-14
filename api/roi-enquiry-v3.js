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
  const m = String(market || "").trim().toLowerCase()
  return (
    m === "us" ||
    m === "usa" ||
    m === "united states" ||
    m === "united states of america"
  )
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

function formatNumber(value) {
  return Number(value || 0).toLocaleString("en-GB")
}

async function buildPdfBuffer(data) {
  return new Promise(async (resolve, reject) => {
    try {
      const doc = new PDFDocument({
        size: getPageSize(data.calculator.market),
        margin: 28,
      })

      const stream = new PassThrough()
      const chunks = []
      stream.on("data", (c) => chunks.push(c))
      stream.on("end", () => resolve(Buffer.concat(chunks)))
      doc.pipe(stream)

      const inputs = data.calculator.inputs || {}
      const assumptions = data.calculator.assumptions || {}
      const outputs = data.calculator.outputs || {}
      const currency = data.calculator.currency || "GBP"

      // === ROI RANGE CALC ===
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

      const lowROI = calcROI({
        reliabilityExposurePct: 3,
        underClaimPct: 3,
        misattributionPct: 1,
        telecomRecoveryPct: 10,
        operationalSavingsPct: 30,
      })

      const highROI = calcROI({
        reliabilityExposurePct: 5,
        underClaimPct: 10,
        misattributionPct: 3,
        telecomRecoveryPct: 25,
        operationalSavingsPct: 70,
      })

      // === HEADER ===
      doc
        .rect(0, 0, doc.page.width, 80)
        .fill("#0B1B3A")

      doc
        .fillColor("#FFFFFF")
        .fontSize(18)
        .text("REE ROI Submission", 28, 20)

      doc
        .fontSize(8)
        .text("Model v3", doc.page.width - 100, 20)

      doc.moveDown(4)

      // === KPI SUMMARY ===
      doc.fillColor("#000000").fontSize(12).text("Summary")

      doc.text(
        `Total Annual Value: ${formatCurrency(
          outputs.totalAnnualValue,
          currency
        )}`
      )

      doc.text(
        `REE Annual Fee: ${formatCurrency(outputs.reeAnnualFee, currency)}`
      )

      doc.text(
        `Net Annual Benefit: ${formatCurrency(
          outputs.netAnnualBenefit,
          currency
        )}`
      )

      doc.text(
        `ROI Multiple: ${formatMultiple(outputs.roiMultiple)} (${formatMultiple(
          lowROI
        )} – ${formatMultiple(highROI)})`
      )

      doc.text(`Payback: ${formatMonths(outputs.paybackMonths)}`)

      doc.moveDown()

      // === BUSINESS INPUTS ===
      doc.fontSize(12).text("Business Inputs (User Supplied)")

      doc.text(`Utility Name: ${data.calculator.utilityName}`)
      doc.text(`Market: ${data.calculator.market}`)
      doc.text(`Number of Meters: ${formatNumber(inputs.numberOfMeters)}`)
      doc.text(
        `Annual Revenue: ${formatCurrency(
          inputs.annualRegulatedRevenue,
          currency
        )}`
      )
      doc.text(
        `Annual Events: ${formatNumber(inputs.annualOutageEvents)}`
      )

      doc.moveDown()

      // === MODEL ASSUMPTIONS ===
      doc.fontSize(12).text("Model Assumptions")

      Object.entries(assumptions).forEach(([k, v]) => {
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

  if (req.method === "OPTIONS") {
    return res.status(200).end()
  }

  try {
    if (!resend) {
      resend = new Resend(process.env.RESEND_API_KEY)
    }

    const data = req.body
    const email = data?.contact?.email

    if (!email) {
      return res.status(400).json({ error: "Missing contact email" })
    }

    const pdfBuffer = await buildPdfBuffer(data)

    await resend.emails.send({
      from: "sales@resonant-grid.com",
      to: email,
      subject: "ROI Summary",
      html: `<p>Your ROI summary is attached.</p>`,
      attachments: [
        {
          filename: "roi-summary.pdf",
          content: pdfBuffer,
        },
      ],
    })

    return res.status(200).json({ success: true })
  } catch (error) {
    console.error(error)
    return res.status(500).json({
      error: error.message || "Server error",
    })
  }
}
