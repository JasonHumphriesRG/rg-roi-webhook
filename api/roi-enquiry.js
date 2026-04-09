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

function sanitizeFilePart(value, fallback = "Unknown_Utility") {
  const cleaned = String(value || fallback)
    .trim()
    .replace(/&/g, "and")
    .replace(/[^a-zA-Z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")

  return cleaned || fallback
}

function buildFileName(utilityName, submittedAt) {
  const utility = sanitizeFilePart(utilityName)
  const dt = new Date(submittedAt || Date.now())

  const yyyy = dt.getFullYear()
  const mm = String(dt.getMonth() + 1).padStart(2, "0")
  const dd = String(dt.getDate()).padStart(2, "0")
  const hh = String(dt.getHours()).padStart(2, "0")
  const mi = String(dt.getMinutes()).padStart(2, "0")
  const ss = String(dt.getSeconds()).padStart(2, "0")

  const stamp = `${yyyy}${mm}${dd}_${hh}${mi}${ss}`

  return `${utility}_Resonant_Grid_Sentinel_Evidence_Engine_${stamp}.pdf`
}

function formatSubmittedAt(value) {
  const dt = new Date(value || Date.now())
  if (Number.isNaN(dt.getTime())) return String(value || "")
  return dt.toLocaleString("en-GB", {
    year: "numeric",
    month: "short",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false,
  })
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

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;")
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

function drawSectionHeader(doc, title, x, y, width) {
  doc
    .roundedRect(x, y, width, 18, 6)
    .fillAndStroke("#F6F9FC", "#E6EDF3")

  doc
    .fillColor("#0B1B3A")
    .font("Helvetica-Bold")
    .fontSize(9)
    .text(title, x + 8, y + 5, { width: width - 16 })

  return y + 24
}

function drawLabelValueRow(doc, label, value, x, y, width, options = {}) {
  const labelWidth = options.labelWidth || 130
  const gap = options.gap || 8
  const valueWidth = width - labelWidth - gap

  const labelText = String(label ?? "—")
  const valueText = String(value ?? "—")

  const labelHeight = doc.heightOfString(labelText, {
    width: labelWidth,
    align: "left",
  })

  const valueHeight = doc.heightOfString(valueText, {
    width: valueWidth,
    align: "left",
  })

  const rowHeight = Math.max(labelHeight, valueHeight, 11)

  doc
    .font("Helvetica-Bold")
    .fontSize(8)
    .fillColor("#6E7A8B")
    .text(labelText, x, y, {
      width: labelWidth,
      align: "left",
      lineGap: 0,
    })

  doc
    .font("Helvetica")
    .fontSize(8)
    .fillColor("#0B1B3A")
    .text(valueText, x + labelWidth + gap, y, {
      width: valueWidth,
      align: "left",
      lineGap: 0,
    })

  return y + rowHeight + 2
}

function fitCoverNote(doc, text, x, y, width, maxBottomY) {
  const noteText = String(text || "—")
  let fontSize = 9
  const minFontSize = 6.5
  const targetHeight = Math.max(24, maxBottomY - y)

  while (fontSize >= minFontSize) {
    doc.font("Helvetica").fontSize(fontSize)
    const h = doc.heightOfString(noteText, {
      width,
      lineGap: 1,
    })
    if (h <= targetHeight) {
      doc.text(noteText, x, y, {
        width,
        lineGap: 1,
      })
      return
    }
    fontSize -= 0.5
  }

  doc.font("Helvetica").fontSize(minFontSize).text(noteText, x, y, {
    width,
    height: targetHeight,
    lineGap: 1,
    ellipsis: true,
  })
}

function buildPdfBuffer(data) {
  return new Promise(async (resolve, reject) => {
    try {
      const pageSize = getPageSize(data?.calculator?.market)

      const doc = new PDFDocument({
        margin: 28,
        size: pageSize,
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
        white: "#FFFFFF",
      }

      const pageWidth =
        doc.page.width - doc.page.margins.left - doc.page.margins.right
      const pageHeight =
        doc.page.height - doc.page.margins.top - doc.page.margins.bottom

      const left = doc.page.margins.left
      const top = doc.page.margins.top
      const gap = 14
      const colWidth = (pageWidth - gap) / 2
      const rightColX = left + colWidth + gap

      const currency = data.calculator?.currency || "GBP"
      const inputs = data.calculator?.inputs || {}
      const assumptions = data.calculator?.assumptions || {}
      const outputs = data.calculator?.outputs || {}

      const reliabilityRevenueUplift =
        Number(inputs.annualRegulatedRevenue || 0) *
        (Number(assumptions.reliabilityExposurePct || 0) / 100) *
        (Number(assumptions.underClaimPct || 0) / 100)

      const misattributionValue =
        Number(inputs.annualOutageEvents || 0) *
        (Number(assumptions.misattributionPct || 0) / 100) *
        Number(inputs.avgCostPerMisclassifiedEvent || 0)

      const telecomRecoveryValue =
        Number(inputs.numberOfMeters || 0) *
        Number(inputs.monthlyCommsCostPerMeter || 0) *
        12 *
        (Number(assumptions.telecomRecoveryPct || 0) / 100)

      const operationalSavings =
        (
          Number(inputs.annualInternalFTECost || 0) +
          Number(inputs.annualConsultantCost || 0) +
          Number(inputs.annualManagementOverhead || 0) +
          Number(inputs.annualAuditRemediationCost || 0)
        ) * (Number(assumptions.operationalSavingsPct || 0) / 100)

      const headerHeight = 82

      doc
        .roundedRect(left, top, pageWidth, headerHeight, 14)
        .fillAndStroke(brand.navy, brand.navy)

      if (logoBuffer) {
        doc.image(logoBuffer, left + 14, top + 18, {
          fit: [105, 32],
          align: "left",
          valign: "center",
        })
      }

      doc
        .fillColor(brand.white)
        .font("Helvetica-Bold")
        .fontSize(18)
        .text("REE ROI Submission", left + 130, top + 12, {
          width: 250,
        })

      doc
        .fillColor(brand.white)
        .font("Helvetica")
        .fontSize(8.5)
        .text(
          "Calculator summary, business inputs, assumptions, outputs, and contact details.",
          left + 130,
          top + 35,
          { width: 270, lineGap: 1 }
        )

      doc
        .fillColor(brand.white)
        .font("Helvetica")
        .fontSize(8)
        .text(
          `Submitted: ${formatSubmittedAt(data.submittedAt)}`,
          left + pageWidth - 180,
          top + 12,
          { width: 165, align: "right" }
        )

      doc
        .font("Helvetica-Bold")
        .fontSize(9)
        .fillColor(brand.white)
        .text(
          data.calculator?.utilityName || "Unknown Utility",
          left + pageWidth - 180,
          top + 32,
          {
            width: 165,
            align: "right",
          }
        )

      doc
        .font("Helvetica")
        .fontSize(8)
        .fillColor(brand.white)
        .text(
          `${data.calculator?.market || "—"} • ${currency}`,
          left + pageWidth - 180,
          top + 46,
          {
            width: 165,
            align: "right",
          }
        )

      let y = top + headerHeight + 10

      const kpiGap = 8
      const kpiWidth = (pageWidth - kpiGap * 2) / 3
      const kpiHeight = 52

      const kpis = [
        {
          title: "Total Annual Value",
          value: formatCurrency(outputs.totalAnnualValue, currency),
          accent: brand.teal,
        },
        {
          title: "REE Annual Fee",
          value: formatCurrency(outputs.reeAnnualFee, currency),
          accent: brand.amber,
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
        {
          title: "Payback",
          value: formatMonths(outputs.paybackMonths),
          accent: brand.teal,
        },
        {
          title: "Breakeven REE Price",
          value: formatCurrency(outputs.breakevenReePrice, currency),
          accent: brand.slate,
        },
      ]

      kpis.forEach((kpi, i) => {
        const row = Math.floor(i / 3)
        const col = i % 3
        const x = left + col * (kpiWidth + kpiGap)
        const cardY = y + row * (kpiHeight + 8)

        doc
          .roundedRect(x, cardY, kpiWidth, kpiHeight, 10)
          .fillAndStroke("#FFFFFF", brand.border)

        doc
          .fillColor(brand.slate)
          .font("Helvetica-Bold")
          .fontSize(7.5)
          .text(kpi.title, x + 10, cardY + 9, {
            width: kpiWidth - 20,
          })

        doc
          .fillColor(brand.navy)
          .font("Helvetica-Bold")
          .fontSize(13)
          .text(kpi.value, x + 10, cardY + 23, {
            width: kpiWidth - 20,
          })

        doc
          .roundedRect(x + 10, cardY + kpiHeight - 9, 30, 3, 2)
          .fill(kpi.accent)
      })

      y += kpiHeight * 2 + 16

      let leftY = y
      leftY = drawSectionHeader(doc, "Contact Details", left, leftY, colWidth)
      leftY = drawLabelValueRow(doc, "Name", data.contact?.name || "—", left, leftY, colWidth)
      leftY = drawLabelValueRow(doc, "Email", data.contact?.email || "—", left, leftY, colWidth)
      leftY = drawLabelValueRow(doc, "Job Title", data.contact?.jobTitle || "—", left, leftY, colWidth)
      leftY += 6

      leftY = drawSectionHeader(doc, "Business Inputs", left, leftY, colWidth)
      leftY = drawLabelValueRow(doc, "Utility Name", data.calculator?.utilityName || "—", left, leftY, colWidth)
      leftY = drawLabelValueRow(doc, "Market", data.calculator?.market || "—", left, leftY, colWidth)
      leftY = drawLabelValueRow(doc, "Currency", currency, left, leftY, colWidth)
      leftY = drawLabelValueRow(doc, "Number of Meters", formatNumber(inputs.numberOfMeters), left, leftY, colWidth)
      leftY = drawLabelValueRow(doc, "Annual Regulated Revenue", formatCurrency(inputs.annualRegulatedRevenue, currency), left, leftY, colWidth)
      leftY = drawLabelValueRow(doc, "Annual Outage / Alert Events", formatNumber(inputs.annualOutageEvents), left, leftY, colWidth)
      leftY = drawLabelValueRow(doc, "Avg. Cost per Misclassified Event", formatCurrency(inputs.avgCostPerMisclassifiedEvent, currency), left, leftY, colWidth)
      leftY = drawLabelValueRow(doc, "Monthly Comms Cost per Meter", formatCurrency2(inputs.monthlyCommsCostPerMeter, currency), left, leftY, colWidth)
      leftY = drawLabelValueRow(doc, "Annual Internal FTE Cost", formatCurrency(inputs.annualInternalFTECost, currency), left, leftY, colWidth)
      leftY = drawLabelValueRow(doc, "Annual Consultant Cost", formatCurrency(inputs.annualConsultantCost, currency), left, leftY, colWidth)
      leftY = drawLabelValueRow(doc, "Annual Management Overhead", formatCurrency(inputs.annualManagementOverhead, currency), left, leftY, colWidth)
      leftY = drawLabelValueRow(doc, "Annual Audit / Remediation Cost", formatCurrency(inputs.annualAuditRemediationCost, currency), left, leftY, colWidth)
      leftY = drawLabelValueRow(doc, "REE Price per Meter / Year", formatCurrency2(inputs.reePricePerMeterPerYear, currency), left, leftY, colWidth)

      let rightY = y
      rightY = drawSectionHeader(doc, "Scenario Assumptions", rightColX, rightY, colWidth)
      rightY = drawLabelValueRow(doc, "Scenario Mode", data.calculator?.scenarioMode || "—", rightColX, rightY, colWidth)
      rightY = drawLabelValueRow(doc, "Scenario Label", data.calculator?.scenarioLabel || "—", rightColX, rightY, colWidth)
      rightY = drawLabelValueRow(doc, "Reliability Exposure", formatPercentWhole(assumptions.reliabilityExposurePct), rightColX, rightY, colWidth)
      rightY = drawLabelValueRow(doc, "Under-Claim", formatPercentWhole(assumptions.underClaimPct), rightColX, rightY, colWidth)
      rightY = drawLabelValueRow(doc, "Misattribution", formatPercentWhole(assumptions.misattributionPct), rightColX, rightY, colWidth)
      rightY = drawLabelValueRow(doc, "Telecom Recovery", formatPercentWhole(assumptions.telecomRecoveryPct), rightColX, rightY, colWidth)
      rightY = drawLabelValueRow(doc, "Operational Savings", formatPercentWhole(assumptions.operationalSavingsPct), rightColX, rightY, colWidth)
      rightY += 6

      rightY = drawSectionHeader(doc, "ROI Outputs", rightColX, rightY, colWidth)
      rightY = drawLabelValueRow(doc, "Total Annual Value", formatCurrency(outputs.totalAnnualValue, currency), rightColX, rightY, colWidth)
      rightY = drawLabelValueRow(doc, "REE Annual Fee", formatCurrency(outputs.reeAnnualFee, currency), rightColX, rightY, colWidth)
      rightY = drawLabelValueRow(doc, "Net Annual Benefit", formatCurrency(outputs.netAnnualBenefit, currency), rightColX, rightY, colWidth)
      rightY = drawLabelValueRow(doc, "ROI Multiple", formatMultiple(outputs.roiMultiple), rightColX, rightY, colWidth)
      rightY = drawLabelValueRow(doc, "Payback", formatMonths(outputs.paybackMonths), rightColX, rightY, colWidth)
      rightY = drawLabelValueRow(doc, "Breakeven REE Price", formatCurrency(outputs.breakevenReePrice, currency), rightColX, rightY, colWidth)
      rightY = drawLabelValueRow(
        doc,
        "REE Share of Value",
        `${Math.round((outputs.reeShareOfValue || 0) * 100)}%`,
        rightColX,
        rightY,
        colWidth
      )

      rightY += 6
      rightY = drawSectionHeader(doc, "Value Build-Up", rightColX, rightY, colWidth)
      rightY = drawLabelValueRow(
        doc,
        "Reliability revenue uplift",
        formatCurrency(reliabilityRevenueUplift, currency),
        rightColX,
        rightY,
        colWidth
      )
      rightY = drawLabelValueRow(
        doc,
        "Misattribution value",
        formatCurrency(misattributionValue, currency),
        rightColX,
        rightY,
        colWidth
      )
      rightY = drawLabelValueRow(
        doc,
        "Telecom recovery",
        formatCurrency(telecomRecoveryValue, currency),
        rightColX,
        rightY,
        colWidth
      )
      rightY = drawLabelValueRow(
        doc,
        "Operational savings",
        formatCurrency(operationalSavings, currency),
        rightColX,
        rightY,
        colWidth
      )

      const noteY = Math.max(leftY, rightY) + 10
      const maxBottomY = top + pageHeight - 6

      const noteHeaderY = drawSectionHeader(doc, "Cover Note", left, noteY, pageWidth)
      doc.fillColor(brand.navy)
      fitCoverNote(
        doc,
        data.contact?.coverNote || "—",
        left,
        noteHeaderY,
        pageWidth,
        maxBottomY
      )

      doc.end()
    } catch (err) {
      reject(err)
    }
  })
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

    const submittedAt = data?.submittedAt || new Date().toISOString()
    const fileName = buildFileName(data?.calculator?.utilityName, submittedAt)

    const pdfBuffer = await buildPdfBuffer({
      ...data,
      submittedAt,
    })

    const driveFile = await uploadToDrive(pdfBuffer, fileName)

    const internalEmail = await resend.emails.send({
      from: "sales@resonant-grid.com",
      to: "WebformROI@resonant-grid.com",
      replyTo: email,
      subject: `ROI Enquiry — ${data?.calculator?.utilityName || "Unknown Utility"}`,
      html: `
        <h2>New ROI Submission</h2>
        <p><strong>Name:</strong> ${escapeHtml(data?.contact?.name || "")}</p>
        <p><strong>Email:</strong> ${escapeHtml(email)}</p>
        <p><strong>Utility:</strong> ${escapeHtml(data?.calculator?.utilityName || "")}</p>
        <p><strong>ROI:</strong> ${escapeHtml(data?.calculator?.outputs?.roiMultiple || "")}x</p>
        <p><strong>Drive File:</strong> ${escapeHtml(driveFile.webViewLink || driveFile.id)}</p>
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
        <p>Thanks ${escapeHtml(data?.contact?.name || "")},</p>
        <p>We’ve received your ROI scenario and will get back to you shortly.</p>
        <p>Your ROI summary PDF is attached for reference.</p>
      `,
      attachments: [
        {
          filename: fileName,
          content: pdfBuffer,
        },
      ],
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
