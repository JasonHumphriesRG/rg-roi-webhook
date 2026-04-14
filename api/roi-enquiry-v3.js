import { Resend } from "resend"
import { google } from "googleapis"
import PDFDocument from "pdfkit"
import { PassThrough, Readable } from "stream"

let resend = null

const PDF_VERSION = "v3"

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

function formatCurrency(value, currency = "GBP", digits = 0) {
  try {
    return new Intl.NumberFormat("en-GB", {
      style: "currency",
      currency,
      minimumFractionDigits: digits,
      maximumFractionDigits: digits,
    }).format(value ?? 0)
  } catch {
    return `${currency} ${Number(value ?? 0).toLocaleString()}`
  }
}

function formatNumber(value, digits = 0) {
  return Number(value ?? 0).toLocaleString("en-GB", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  })
}

function formatPercentWhole(value, digits = 0) {
  return `${Number(value ?? 0).toFixed(digits)}%`
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
    .replace(/\"/g, "&quot;")
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

function calculateModel(data) {
  const calculator = data?.calculator || {}
  const inputs = calculator.inputs || {}
  const assumptions = calculator.assumptions || {}

  const numberOfMeters = Number(inputs.numberOfMeters || 0)
  const annualRegulatedRevenue = Number(inputs.annualRegulatedRevenue || 0)
  const outageEventsPerMeterPerYear = Number(
    inputs.outageEventsPerMeterPerYear || 0
  )
  const avgCostPerMisclassifiedEvent = Number(
    inputs.avgCostPerMisclassifiedEvent || 0
  )
  const regulatoryReportingCostPerMeter = Number(
    inputs.regulatoryReportingCostPerMeter || 0
  )
  const monthlyCommsCostPerMeter = Number(inputs.monthlyCommsCostPerMeter || 0)
  const reePricePerMeterPerYear = Number(inputs.reePricePerMeterPerYear || 0)

  const reliabilityExposurePct = Number(
    assumptions.reliabilityExposurePct || 0
  )
  const underClaimPct = Number(assumptions.underClaimPct || 0)
  const misattributionPct = Number(assumptions.misattributionPct || 0)
  const operationalSavingsPct = Number(
    assumptions.operationalSavingsPct || 0
  )
  const telecomRecoveryPct = Number(assumptions.telecomRecoveryPct || 0)

  const reliabilityRevenueUplift =
    annualRegulatedRevenue *
    (reliabilityExposurePct / 100) *
    (underClaimPct / 100)

  const totalEvents = numberOfMeters * outageEventsPerMeterPerYear
  const misattributionEvents = totalEvents * (misattributionPct / 100)
  const misattributionValue =
    misattributionEvents * avgCostPerMisclassifiedEvent

  const operationalSavings =
    numberOfMeters *
    regulatoryReportingCostPerMeter *
    (operationalSavingsPct / 100)

  const telecomRecoveryValue =
    numberOfMeters *
    monthlyCommsCostPerMeter *
    12 *
    (telecomRecoveryPct / 100)

  const totalAnnualValue =
    reliabilityRevenueUplift +
    misattributionValue +
    operationalSavings +
    telecomRecoveryValue

  const reeAnnualFee = numberOfMeters * reePricePerMeterPerYear
  const netAnnualBenefit = totalAnnualValue - reeAnnualFee
  const roiMultiple = reeAnnualFee > 0 ? totalAnnualValue / reeAnnualFee : 0
  const paybackMonths =
    totalAnnualValue > 0 ? reeAnnualFee / (totalAnnualValue / 12) : Infinity
  const breakevenReePrice =
    numberOfMeters > 0 ? totalAnnualValue / numberOfMeters : 0
  const reeShareOfValue =
    totalAnnualValue > 0 ? reeAnnualFee / totalAnnualValue : 0
  const valuePerMeter = numberOfMeters > 0 ? totalAnnualValue / numberOfMeters : 0

  return {
    totalEvents,
    misattributionEvents,
    reliabilityRevenueUplift,
    misattributionValue,
    operationalSavings,
    telecomRecoveryValue,
    totalAnnualValue,
    reeAnnualFee,
    netAnnualBenefit,
    roiMultiple,
    paybackMonths: isFinite(paybackMonths) ? paybackMonths : null,
    breakevenReePrice,
    reeShareOfValue,
    valuePerMeter,
  }
}

function calculateRoiRange(data) {
  const calculator = data?.calculator || {}
  const assumptions = calculator.assumptions || {}

  const lowData = {
    ...data,
    calculator: {
      ...calculator,
      assumptions: {
        ...assumptions,
        reliabilityExposurePct:
          Number(assumptions.reliabilityExposurePct || 0) * 0.6,
        underClaimPct: Number(assumptions.underClaimPct || 0) * 0.6,
        misattributionPct: Number(assumptions.misattributionPct || 0) * 0.5,
        telecomRecoveryPct:
          Number(assumptions.telecomRecoveryPct || 0) * 0.5,
        operationalSavingsPct:
          Number(assumptions.operationalSavingsPct || 0) * 0.6,
      },
    },
  }

  const highData = {
    ...data,
    calculator: {
      ...calculator,
      assumptions: {
        ...assumptions,
        reliabilityExposurePct: Number(
          assumptions.reliabilityExposurePct || 0
        ),
        underClaimPct: Number(assumptions.underClaimPct || 0) * 2,
        misattributionPct: Number(assumptions.misattributionPct || 0) * 1.5,
        telecomRecoveryPct:
          Number(assumptions.telecomRecoveryPct || 0) * 1.25,
        operationalSavingsPct:
          Number(assumptions.operationalSavingsPct || 0) * 1.4,
      },
    },
  }

  return {
    low: calculateModel(lowData),
    high: calculateModel(highData),
  }
}

function drawSectionHeader(doc, title, x, y, width) {
  doc.roundedRect(x, y, width, 18, 6).fillAndStroke("#F6F9FC", "#E6EDF3")

  doc
    .fillColor("#0B1B3A")
    .font("Helvetica-Bold")
    .fontSize(9)
    .text(title, x + 8, y + 5, { width: width - 16 })

  return y + 24
}

function drawLabelValueRow(doc, label, value, x, y, width, options = {}) {
  const labelWidth = options.labelWidth || 118
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
      const outputs = calculateModel(data)
      const roiRange = calculateRoiRange(data)

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
          "Calculator summary, business inputs, hardwired assumptions, outputs, and contact details.",
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
          `${data.calculator?.market || "—"} • ${currency} • ${data.calculator?.modelVersion || PDF_VERSION}`,
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
          title: "ROI Range",
          value: `${formatMultiple(roiRange.low.roiMultiple)} – ${formatMultiple(roiRange.high.roiMultiple)}`,
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

        doc.roundedRect(x + 10, cardY + kpiHeight - 9, 30, 3, 2).fill(kpi.accent)
      })

      y += kpiHeight * 2 + 16

      let leftY = y
      leftY = drawSectionHeader(doc, "Contact Details", left, leftY, colWidth)
      leftY = drawLabelValueRow(
        doc,
        "Name",
        data.contact?.name || "—",
        left,
        leftY,
        colWidth
      )
      leftY = drawLabelValueRow(
        doc,
        "Email",
        data.contact?.email || "—",
        left,
        leftY,
        colWidth
      )
      leftY = drawLabelValueRow(
        doc,
        "Job Title",
        data.contact?.jobTitle || "—",
        left,
        leftY,
        colWidth
      )
      leftY += 6

      leftY = drawSectionHeader(
        doc,
        "Business Inputs (User Supplied)",
        left,
        leftY,
        colWidth
      )
      leftY = drawLabelValueRow(
        doc,
        "Utility Name",
        data.calculator?.utilityName || "—",
        left,
        leftY,
        colWidth
      )
      leftY = drawLabelValueRow(
        doc,
        "Market",
        data.calculator?.market || "—",
        left,
        leftY,
        colWidth
      )
      leftY = drawLabelValueRow(doc, "Currency", currency, left, leftY, colWidth)
      leftY = drawLabelValueRow(
        doc,
        "Number of Meters",
        formatNumber(inputs.numberOfMeters),
        left,
        leftY,
        colWidth
      )
      leftY = drawLabelValueRow(
        doc,
        "Annual Regulated Revenue",
        formatCurrency(inputs.annualRegulatedRevenue, currency),
        left,
        leftY,
        colWidth
      )

      let rightY = y
      rightY = drawSectionHeader(
        doc,
        "Model Assumptions",
        rightColX,
        rightY,
        colWidth
      )
      rightY = drawLabelValueRow(
        doc,
        "Outage Events per Meter / Year",
        formatNumber(inputs.outageEventsPerMeterPerYear, 2),
        rightColX,
        rightY,
        colWidth
      )
      rightY = drawLabelValueRow(
        doc,
        "Cost per Misclassified Event",
        formatCurrency(inputs.avgCostPerMisclassifiedEvent, currency, 2),
        rightColX,
        rightY,
        colWidth
      )
      rightY = drawLabelValueRow(
        doc,
        "Regulatory Reporting Cost / Meter",
        formatCurrency(inputs.regulatoryReportingCostPerMeter, currency, 2),
        rightColX,
        rightY,
        colWidth
      )
      rightY = drawLabelValueRow(
        doc,
        "Comms Cost per Meter / Month",
        formatCurrency(inputs.monthlyCommsCostPerMeter, currency, 2),
        rightColX,
        rightY,
        colWidth
      )
      rightY = drawLabelValueRow(
        doc,
        "REE Price per Meter / Year",
        formatCurrency(inputs.reePricePerMeterPerYear, currency, 2),
        rightColX,
        rightY,
        colWidth
      )
      rightY = drawLabelValueRow(
        doc,
        "Revenue linked to reliability",
        formatPercentWhole(assumptions.reliabilityExposurePct, 1),
        rightColX,
        rightY,
        colWidth
      )
      rightY = drawLabelValueRow(
        doc,
        "Under-claim due to uncertainty",
        formatPercentWhole(assumptions.underClaimPct, 1),
        rightColX,
        rightY,
        colWidth
      )
      rightY = drawLabelValueRow(
        doc,
        "Misattributed from comms",
        formatPercentWhole(assumptions.misattributionPct, 1),
        rightColX,
        rightY,
        colWidth
      )
      rightY = drawLabelValueRow(
        doc,
        "Cost reduction achievable",
        formatPercentWhole(assumptions.operationalSavingsPct, 1),
        rightColX,
        rightY,
        colWidth
      )
      rightY = drawLabelValueRow(
        doc,
        "Recoverable from telecom provider",
        formatPercentWhole(assumptions.telecomRecoveryPct, 1),
        rightColX,
        rightY,
        colWidth
      )
      rightY += 6

      rightY = drawSectionHeader(doc, "ROI Outputs", rightColX, rightY, colWidth)
      rightY = drawLabelValueRow(
        doc,
        "Total Events",
        formatNumber(outputs.totalEvents),
        rightColX,
        rightY,
        colWidth
      )
      rightY = drawLabelValueRow(
        doc,
        "Misattribution Events",
        formatNumber(outputs.misattributionEvents),
        rightColX,
        rightY,
        colWidth
      )
      rightY = drawLabelValueRow(
        doc,
        "Total Annual Value",
        formatCurrency(outputs.totalAnnualValue, currency),
        rightColX,
        rightY,
        colWidth
      )
      rightY = drawLabelValueRow(
        doc,
        "REE Annual Fee",
        formatCurrency(outputs.reeAnnualFee, currency),
        rightColX,
        rightY,
        colWidth
      )
      rightY = drawLabelValueRow(
        doc,
        "Net Annual Benefit",
        formatCurrency(outputs.netAnnualBenefit, currency),
        rightColX,
        rightY,
        colWidth
      )
      rightY = drawLabelValueRow(
        doc,
        "ROI Multiple",
        formatMultiple(outputs.roiMultiple),
        rightColX,
        rightY,
        colWidth
      )
      rightY = drawLabelValueRow(
        doc,
        "Payback",
        formatMonths(outputs.paybackMonths),
        rightColX,
        rightY,
        colWidth
      )
      rightY = drawLabelValueRow(
        doc,
        "ROI Range",
        `${formatMultiple(roiRange.low.roiMultiple)} – ${formatMultiple(roiRange.high.roiMultiple)}`,
        rightColX,
        rightY,
        colWidth
      )
      rightY = drawLabelValueRow(
        doc,
        "REE Share of Value",
        formatPercentWhole((outputs.reeShareOfValue || 0) * 100, 1),
        rightColX,
        rightY,
        colWidth
      )
      rightY = drawLabelValueRow(
        doc,
        "Value per Meter",
        formatCurrency(outputs.valuePerMeter, currency, 2),
        rightColX,
        rightY,
        colWidth
      )

      rightY += 6
      rightY = drawSectionHeader(
        doc,
        "Value Build-Up",
        rightColX,
        rightY,
        colWidth
      )
      rightY = drawLabelValueRow(
        doc,
        "Reliability revenue uplift",
        formatCurrency(outputs.reliabilityRevenueUplift, currency),
        rightColX,
        rightY,
        colWidth
      )
      rightY = drawLabelValueRow(
        doc,
        "Misattribution value",
        formatCurrency(outputs.misattributionValue, currency),
        rightColX,
        rightY,
        colWidth
      )
      rightY = drawLabelValueRow(
        doc,
        "Operational savings",
        formatCurrency(outputs.operationalSavings, currency),
        rightColX,
        rightY,
        colWidth
      )
      rightY = drawLabelValueRow(
        doc,
        "Telecom recovery",
        formatCurrency(outputs.telecomRecoveryValue, currency),
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
      return res.status(500).json({
        error: "Missing GOOGLE_SERVICE_ACCOUNT_JSON",
      })
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

    const outputs = calculateModel(data)
    const submittedAt = data?.submittedAt || new Date().toISOString()
    const fileName = buildFileName(data?.calculator?.utilityName, submittedAt)

    const payload = {
      ...data,
      submittedAt,
      calculator: {
        ...(data?.calculator || {}),
        outputs,
      },
    }

    const pdfBuffer = await buildPdfBuffer(payload)
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
        <p><strong>Total Annual Value:</strong> ${escapeHtml(formatCurrency(outputs.totalAnnualValue, data?.calculator?.currency || "GBP"))}</p>
        <p><strong>ROI:</strong> ${escapeHtml(formatMultiple(outputs.roiMultiple))}</p>
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
        <div style="font-family: Arial, sans-serif; color: #0B1B3A;">
          <div style="background-color:#0B1B3A; padding:16px; text-align:left;">
            ${
              process.env.PDF_LOGO_URL
                ? `<img src="${process.env.PDF_LOGO_URL}" alt="Resonant Grid" style="height:32px;" />`
                : `<span style="color:#FFFFFF; font-weight:bold;">Resonant Grid</span>`
            }
          </div>

          <div style="padding:20px; font-size:14px; line-height:1.5;">
            <p>Thanks ${escapeHtml(data?.contact?.name || "")},</p>

            <p>
              We’ve received your ROI scenario and will get back to you shortly.
              Your ROI summary PDF is attached for reference.
            </p>

            <p>
              If you'd like to get in touch directly with us please email
              <strong>Michael Jary</strong>
              (<a href="mailto:michael@resonant-grid.com">michael@resonant-grid.com</a>)
              and we'll come back to you shortly.
            </p>

            <p style="margin-top:24px;">
              With thanks for your interest,<br/>
              <strong>from the Resonant Grid Team</strong>
            </p>
          </div>
        </div>
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
      outputs,
    })
  } catch (error) {
    console.error(error)
    return res.status(500).json({
      error: error.message || "Server error",
    })
  }
}
