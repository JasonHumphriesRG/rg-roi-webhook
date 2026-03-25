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

function formatCurrency(value, currency = "GBP") {
  try {
    return new Intl.NumberFormat("en-GB", {
      style: "currency",
      currency,
      maximumFractionDigits: 0,
    }).format(value ?? 0)
  } catch {
    return `${currency} ${Number(value ?? 0).toLocaleString()}`
  }
}

function formatCurrency2(value, currency = "GBP") {
  try {
    return new Intl.NumberFormat("en-GB", {
      style: "currency",
      currency,
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    }).format(value ?? 0)
  } catch {
    return `${currency} ${Number(value ?? 0).toFixed(2)}`
  }
}

function formatNumber(value) {
  return Number(value ?? 0).toLocaleString("en-GB")
}

function formatPercentWhole(value) {
  return `${Math.round(Number(value ?? 0))}%`
}

function formatMultiple(value) {
  return `${Number(value ?? 0).toFixed(1)}x`
}

function formatMonths(value) {
  if (value == null || !isFinite(value)) return "—"
  return `${Number(value).toFixed(1)} months`
}

async function loadLogoBuffer() {
  if (!process.env.PDF_LOGO_URL) return null

  const response = await fetch(process.env.PDF_LOGO_URL)
  if (!response.ok) {
    throw new Error(`Failed to load logo from PDF_LOGO_URL: ${response.status}`)
  }
  const arrayBuffer = await response.arrayBuffer()
  return Buffer.from(arrayBuffer)
}

function drawLabelValueRow(doc, label, value, x, y, width, options = {}) {
  const labelWidth = options.labelWidth || 220
  const valueWidth = width - labelWidth
  doc
    .font("Helvetica-Bold")
    .fontSize(10)
    .fillColor("#6E7A8B")
    .text(label, x, y, { width: labelWidth, continued: false })

  doc
    .font("Helvetica")
    .fontSize(10)
    .fillColor("#0B1B3A")
    .text(String(value ?? "—"), x + labelWidth, y, { width: valueWidth })

  return doc.y
}

function drawSectionHeader(doc, title, x, y, width) {
  doc
    .roundedRect(x, y, width, 22, 8)
    .fillAndStroke("#F6F9FC", "#E6EDF3")

  doc
    .fillColor("#0B1B3A")
    .font("Helvetica-Bold")
    .fontSize(11)
    .text(title, x + 10, y + 6, { width: width - 20 })

  return y + 30
}

async function buildPdfBuffer(data) {
  return new Promise(async (resolve, reject) => {
    try {
      const doc = new PDFDocument({
        margin: 40,
        size: "A4",
      })

      const stream = new PassThrough()
      const chunks = []

      stream.on("data", (chunk) => chunks.push(chunk))
      stream.on("end", () => resolve(Buffer.concat(chunks)))
      stream.on("error", reject)

      doc.pipe(stream)

      const logoBuffer = await loadLogoBuffer().catch(() => null)

      const brand = {
        navy: "#0B1B3A",
        teal: "#00B5A4",
        green: "#3ED37A",
        amber: "#F4A300",
        slate: "#6E7A8B",
        cloud: "#F6F9FC",
        border: "#E6EDF3",
      }

      const pageWidth =
        doc.page.width - doc.page.margins.left - doc.page.margins.right
      const left = doc.page.margins.left
      const rightColX = left + pageWidth / 2 + 10
      const colWidth = pageWidth / 2 - 10

      // Header card
      doc
        .roundedRect(left, 35, pageWidth, 90, 16)
        .fillAndStroke("#FFFFFF", brand.border)

      if (logoBuffer) {
        doc.image(logoBuffer, left + 18, 52, { fit: [110, 40] })
      }

      doc
        .fillColor(brand.navy)
        .font("Helvetica-Bold")
        .fontSize(22)
        .text("REE ROI Submission", left + 150, 52, { width: 260 })

      doc
        .fillColor(brand.slate)
        .font("Helvetica")
        .fontSize(11)
        .text(
          "Calculator summary, submission details, business inputs, assumptions, and ROI outputs.",
          left + 150,
          82,
          { width: 330 }
        )

      doc
        .font("Helvetica")
        .fontSize(10)
        .fillColor(brand.slate)
        .text(
          `Submitted: ${data.submittedAt || new Date().toISOString()}`,
          left + pageWidth - 180,
          52,
          { width: 160, align: "right" }
        )

      let y = 145

      // Summary KPIs strip
      const kpiY = y
      const kpiWidth = (pageWidth - 30) / 3
      const outputs = data.calculator?.outputs || {}
      const currency = data.calculator?.currency || "GBP"

      const kpis = [
        {
          title: "Total Annual Value",
          value: formatCurrency(outputs.totalAnnualValue, currency),
          accent: brand.teal,
        },
        {
          title: "Net Annual Benefit",
          value: formatCurrency(outputs.netAnnualBenefit, currency),
          accent: outputs.netAnnualBenefit >= 0 ? brand.green : brand.amber,
        },
        {
          title: "ROI Multiple",
          value: formatMultiple(outputs.roiMultiple),
          accent: brand.navy,
        },
      ]

      kpis.forEach((kpi, i) => {
        const x = left + i * (kpiWidth + 15)
        doc.roundedRect(x, kpiY, kpiWidth, 72, 14).fillAndStroke("#FFFFFF", brand.border)
        doc.fillColor(brand.slate).font("Helvetica-Bold").fontSize(9).text(kpi.title, x + 12, kpiY + 12, {
          width: kpiWidth - 24,
        })
        doc.fillColor(brand.navy).font("Helvetica-Bold").fontSize(19).text(kpi.value, x + 12, kpiY + 30, {
          width: kpiWidth - 24,
        })
        doc.roundedRect(x + 12, kpiY + 58, 32, 4, 2).fill(kpi.accent)
      })

      y = kpiY + 92

      // Left column
      let leftY = y
      leftY = drawSectionHeader(doc, "Contact Details", left, leftY, colWidth)
      leftY = drawLabelValueRow(doc, "Name", data.contact?.name || "—", left, leftY, colWidth)
      leftY += 4
      leftY = drawLabelValueRow(doc, "Email", data.contact?.email || "—", left, leftY, colWidth)
      leftY += 4
      leftY = drawLabelValueRow(doc, "Job Title", data.contact?.jobTitle || "—", left, leftY, colWidth)
      leftY += 14

      leftY = drawSectionHeader(doc, "Business Inputs", left, leftY, colWidth)
      const inputs = data.calculator?.inputs || {}

      leftY = drawLabelValueRow(doc, "Utility Name", data.calculator?.utilityName || "—", left, leftY, colWidth)
      leftY += 4
      leftY = drawLabelValueRow(doc, "Market", data.calculator?.market || "—", left, leftY, colWidth)
      leftY += 4
      leftY = drawLabelValueRow(doc, "Number of Meters", formatNumber(inputs.numberOfMeters), left, leftY, colWidth)
      leftY += 4
      leftY = drawLabelValueRow(doc, "Annual Regulated Revenue", formatCurrency(inputs.annualRegulatedRevenue, currency), left, leftY, colWidth)
      leftY += 4
      leftY = drawLabelValueRow(doc, "Annual Outage / Alert Events", formatNumber(inputs.annualOutageEvents), left, leftY, colWidth)
      leftY += 4
      leftY = drawLabelValueRow(doc, "Avg. Cost per Misclassified Event", formatCurrency(inputs.avgCostPerMisclassifiedEvent, currency), left, leftY, colWidth)
      leftY += 4
      leftY = drawLabelValueRow(doc, "Monthly Comms Cost per Meter", formatCurrency2(inputs.monthlyCommsCostPerMeter, currency), left, leftY, colWidth)
      leftY += 4
      leftY = drawLabelValueRow(doc, "Annual Internal FTE Cost", formatCurrency(inputs.annualInternalFTECost, currency), left, leftY, colWidth)
      leftY += 4
      leftY = drawLabelValueRow(doc, "Annual Consultant Cost", formatCurrency(inputs.annualConsultantCost, currency), left, leftY, colWidth)
      leftY += 4
      leftY = drawLabelValueRow(doc, "Annual Management Overhead", formatCurrency(inputs.annualManagementOverhead, currency), left, leftY, colWidth)
      leftY += 4
      leftY = drawLabelValueRow(doc, "Annual Audit / Remediation Cost", formatCurrency(inputs.annualAuditRemediationCost, currency), left, leftY, colWidth)
      leftY += 4
      leftY = drawLabelValueRow(doc, "REE Price per Meter / Year", formatCurrency2(inputs.reePricePerMeterPerYear, currency), left, leftY, colWidth)

      // Right column
      let rightY = y
      rightY = drawSectionHeader(doc, "Scenario Assumptions", rightColX, rightY, colWidth)
      const assumptions = data.calculator?.assumptions || {}

      rightY = drawLabelValueRow(doc, "Scenario Mode", data.calculator?.scenarioMode || "—", rightColX, rightY, colWidth)
      rightY += 4
      rightY = drawLabelValueRow(doc, "Scenario Label", data.calculator?.scenarioLabel || "—", rightColX, rightY, colWidth)
      rightY += 4
      rightY = drawLabelValueRow(doc, "Reliability Exposure", formatWholePercent(assumptions.reliabilityExposurePct), rightColX, rightY, colWidth)
      rightY += 4
      rightY = drawLabelValueRow(doc, "Under-Claim", formatWholePercent(assumptions.underClaimPct), rightColX, rightY, colWidth)
      rightY += 4
      rightY = drawLabelValueRow(doc, "Misattribution", formatWholePercent(assumptions.misattributionPct), rightColX, rightY, colWidth)
      rightY += 4
      rightY = drawLabelValueRow(doc, "Telecom Recovery", formatWholePercent(assumptions.telecomRecoveryPct), rightColX, rightY, colWidth)
      rightY += 4
      rightY = drawLabelValueRow(doc, "Operational Savings", formatWholePercent(assumptions.operationalSavingsPct), rightColX, rightY, colWidth)
      rightY += 14

      rightY = drawSectionHeader(doc, "ROI Outputs", rightColX, rightY, colWidth)
      rightY = drawLabelValueRow(doc, "Total Annual Value", formatCurrency(outputs.totalAnnualValue, currency), rightColX, rightY, colWidth)
      rightY += 4
      rightY = drawLabelValueRow(doc, "REE Annual Fee", formatCurrency(outputs.reeAnnualFee, currency), rightColX, rightY, colWidth)
      rightY += 4
      rightY = drawLabelValueRow(doc, "Net Annual Benefit", formatCurrency(outputs.netAnnualBenefit, currency), rightColX, rightY, colWidth)
      rightY += 4
      rightY = drawLabelValueRow(doc, "ROI Multiple", formatMultiple(outputs.roiMultiple), rightColX, rightY, colWidth)
      rightY += 4
      rightY = drawLabelValueRow(doc, "Payback", formatMonths(outputs.paybackMonths), rightColX, rightY, colWidth)
      rightY += 4
      rightY = drawLabelValueRow(doc, "Breakeven REE Price", formatCurrency(outputs.breakevenReePrice, currency), rightColX, rightY, colWidth)
      rightY += 4
      rightY = drawLabelValueRow(
        doc,
        "REE Share of Value",
        `${Math.round((outputs.reeShareOfValue || 0) * 100)}%`,
        rightColX,
        rightY,
        colWidth
      )

      const noteY = Math.max(leftY, rightY) + 18
      drawSectionHeader(doc, "Cover Note", left, noteY, pageWidth)
      doc
        .fillColor("#0B1B3A")
        .font("Helvetica")
        .fontSize(10)
        .text(data.contact?.coverNote || "—", left, noteY + 34, {
          width: pageWidth,
          lineGap: 4,
        })

      doc.end()
    } catch (err) {
      reject(err)
    }
  })
}

async function loadLogoBuffer() {
  if (!process.env.PDF_LOGO_URL) return null
  const response = await fetch(process.env.PDF_LOGO_URL)
  if (!response.ok) throw new Error(`Failed to load logo from PDF_LOGO_URL`)
  const arrayBuffer = await response.arrayBuffer()
  return Buffer.from(arrayBuffer)
}

async function uploadToDrive(pdfBuffer, fileName) {
  if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    throw new Error("Missing GOOGLE_SERVICE_ACCOUNT_JSON")
  }

  if (!process.env.GOOGLE_DRIVE_FOLDER_ID) {
    throw new Error("Missing GOOGLE_DRIVE_FOLDER_ID")
  }

  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON)

  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/drive.file"],
  })

  const drive = google.drive({ version: "v3", auth })

  const response = await drive.files.create({
    requestBody: {
      name: fileName,
      parents: [process.env.GOOGLE_DRIVE_FOLDER_ID],
    },
    media: {
      mimeType: "application/pdf",
      body: Readable.from(pdfBuffer),
    },
    fields: "id,name,webViewLink",
    supportsAllDrives: true,
  })

  return response.data
}

export default async function handler(req, res) {
  setCors(res)

  if (req.method === "OPTIONS") {
    return res.status(200).end()
  }

  try {
    if (!process.env.RESEND_API_KEY) {
      return res.status(500).json({ error: "Missing RESEND_API_KEY" })
    }

    if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
      return res.status(500).json({ error: "Missing GOOGLE_SERVICE_ACCOUNT_JSON" })
    }

    if (!process.env.GOOGLE_DRIVE_FOLDER_ID) {
      return res.status(500).json({ error: "Missing GOOGLE_DRIVE_FOLDER_ID" })
    }

    if (!resend) {
      resend = new Resend(process.env.RESEND_API_KEY)
    }

    if (req.method !== "POST") {
      return res.status(405).json({ error: "Method not allowed" })
    }

    const data = req.body
    const email = data?.contact?.email

    if (!email) {
      return res.status(400).json({ error: "Missing contact email" })
    }

    const timestamp = new Date().toISOString()
    const safeEmail = String(email).replace(/[^a-zA-Z0-9]/g, "_")
    const fileName = `${timestamp}__${safeEmail}__roi.pdf`

    const pdfBuffer = await buildPdfBuffer(data)
    const driveFile = await uploadToDrive(pdfBuffer, fileName)

    const internalEmail = await resend.emails.send({
      from: "sales@resonant-grid.com",
      to: "jason@resonant-grid.com",
      replyTo: email,
      subject: `ROI Enquiry — ${data?.calculator?.utilityName || "Unknown Utility"}`,
      html: `
        <h2>New ROI Submission</h2>
        <p><strong>Name:</strong> ${data?.contact?.name || ""}</p>
        <p><strong>Email:</strong> ${email}</p>
        <p><strong>Utility:</strong> ${data?.calculator?.utilityName || ""}</p>
        <p><strong>ROI:</strong> ${data?.calculator?.outputs?.roiMultiple || ""}x</p>
        <p><strong>Drive File:</strong> ${driveFile.webViewLink || driveFile.id}</p>
      `,
    })

    if (internalEmail?.error) {
      throw new Error(
        `Internal email failed: ${
          internalEmail.error.message || JSON.stringify(internalEmail.error)
        }`
      )
    }

    const confirmationEmail = await resend.emails.send({
      from: "sales@resonant-grid.com",
      to: email,
      subject: "We’ve received your ROI submission",
      html: `
        <p>Thanks ${data?.contact?.name || ""},</p>
        <p>We’ve received your ROI scenario and will get back to you shortly.</p>
      `,
    })

    if (confirmationEmail?.error) {
      throw new Error(
        `Confirmation email failed: ${
          confirmationEmail.error.message ||
          JSON.stringify(confirmationEmail.error)
        }`
      )
    }

    return res.status(200).json({
      success: true,
      fileId: driveFile.id,
      fileName: driveFile.name,
      driveUrl: driveFile.webViewLink,
    })
  } catch (error) {
    console.error(error)
    return res.status(500).json({
      error: error.message || "Server error",
    })
  }
}
