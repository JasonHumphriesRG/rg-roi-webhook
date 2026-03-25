import { Resend } from "resend"

let resend = null

export default async function handler(req, res) {
  if (!process.env.RESEND_API_KEY) {
    return res.status(500).json({ error: "Missing RESEND_API_KEY" })
  }

  if (!resend) {
    resend = new Resend(process.env.RESEND_API_KEY)
  }

  return res.status(200).json({ ok: true, version: "probe-2" })
}
