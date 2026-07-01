import nodemailer from 'nodemailer'

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' })
  }

  const { to, displayName, lines, delId } = req.body

  if (!to || !lines?.length) {
    return res.status(400).json({ error: 'Faltan datos' })
  }

  const transporter = nodemailer.createTransport({
    host: 'smtp.office365.com',
    port: 587,
    secure: false,
    auth: {
      user: process.env.SMTP_USER,
      pass: process.env.SMTP_PASS,
    },
    tls: { ciphers: 'SSLv3' }
  })

  const linesList = lines
    .map(l => `• ${l.name} – Talle ${l.talle} (x${l.qty})`)
    .join('\n')

  const html = `
    <div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto">
      <div style="background:#121212;padding:24px;text-align:center">
        <div style="color:#FFD200;font-weight:bold;font-size:18px;letter-spacing:1px">INDUMENTARIA PEÑAROL</div>
      </div>
      <div style="padding:28px 24px;background:#fff;border:1px solid #e0e0e0">
        <p style="margin:0 0 12px;font-size:15px">Hola <strong>${displayName}</strong>,</p>
        <p style="margin:0 0 18px;font-size:15px">
          Tenés una entrega de indumentaria pendiente de confirmación desde el depósito.
        </p>
        <div style="background:#f8f8f8;border:1px solid #e0e0e0;border-radius:8px;padding:16px;margin-bottom:20px">
          <div style="font-weight:bold;font-size:12px;color:#888;letter-spacing:.05em;margin-bottom:10px">DETALLE DE ENTREGA #${delId}</div>
          ${lines.map(l => `
            <div style="display:flex;justify-content:space-between;padding:6px 0;border-bottom:1px solid #eee;font-size:14px">
              <span>${l.name} – Talle ${l.talle}</span>
              <span style="font-weight:bold">x${l.qty}</span>
            </div>`).join('')}
        </div>
        <a href="https://inventario-eds1891.vercel.app" style="display:inline-block;background:#FFD200;color:#121212;font-weight:bold;padding:12px 28px;border-radius:6px;text-decoration:none;font-size:14px">
          Aceptar o rechazar entrega
        </a>
      </div>
      <div style="padding:16px 24px;background:#f0f0f0;font-size:12px;color:#888;text-align:center">
        Depósito Peñarol &nbsp;·&nbsp; compras@capenarol.com.uy
      </div>
    </div>
  `

  await transporter.sendMail({
    from: `"Depósito Peñarol" <${process.env.SMTP_USER}>`,
    to,
    subject: `Entrega #${delId} pendiente de confirmación – Depósito Peñarol`,
    text: `Hola ${displayName},\n\nTenés una entrega de indumentaria pendiente de confirmación.\n\nDetalle:\n${linesList}\n\nIngresá a https://inventario-eds1891.vercel.app para aceptar o rechazar la entrega.\n\nDepósito Peñarol`,
    html,
  })

  res.status(200).json({ ok: true })
}
