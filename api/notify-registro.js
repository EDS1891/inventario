import nodemailer from 'nodemailer'

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*')
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS')
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type')
  if (req.method === 'OPTIONS') return res.status(200).end()

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' })
  }

  const { displayName, email, telefono, cargo, categoria, division } = req.body

  if (!displayName || !email) {
    return res.status(400).json({ error: 'Faltan datos' })
  }

  try {
    const transporter = nodemailer.createTransport({
      host: 'smtp.office365.com',
      port: 587,
      secure: false,
      requireTLS: true,
      auth: {
        user: process.env.SMTP_USER,
        pass: process.env.SMTP_PASS,
      },
      tls: { rejectUnauthorized: false }
    })

    const sectorLine = categoria ? `<tr><td style="color:#888;padding:4px 8px 4px 0">Sector</td><td style="padding:4px 0">${categoria}${division ? ` / ${division}` : ''}</td></tr>` : ''

    const html = `
      <div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto">
        <div style="background:#121212;padding:24px;text-align:center">
          <div style="color:#FFD200;font-weight:bold;font-size:18px;letter-spacing:1px">INDUMENTARIA PEÑAROL</div>
        </div>
        <div style="padding:28px 24px;background:#fff;border:1px solid #e0e0e0">
          <p style="margin:0 0 12px;font-size:15px">Nuevo usuario registrado en el sistema:</p>
          <div style="background:#f8f8f8;border:1px solid #e0e0e0;border-radius:8px;padding:16px;margin-bottom:20px">
            <table style="font-size:14px;width:100%;border-collapse:collapse">
              <tr><td style="color:#888;padding:4px 8px 4px 0">Nombre</td><td style="padding:4px 0;font-weight:bold">${displayName}</td></tr>
              <tr><td style="color:#888;padding:4px 8px 4px 0">Email</td><td style="padding:4px 0">${email}</td></tr>
              <tr><td style="color:#888;padding:4px 8px 4px 0">Teléfono</td><td style="padding:4px 0">${telefono || '—'}</td></tr>
              <tr><td style="color:#888;padding:4px 8px 4px 0">Cargo</td><td style="padding:4px 0">${cargo || '—'}</td></tr>
              ${sectorLine}
            </table>
          </div>
          <p style="font-size:13px;color:#555;margin:0">El usuario ya puede ingresar al sistema con su correo y contraseña.<br/>Podés cambiar su rol desde <strong>Gestión de usuarios</strong>.</p>
          <div style="margin-top:20px">
            <a href="https://indumentaria-cap.vercel.app" style="display:inline-block;background:#FFD200;color:#121212;font-weight:bold;padding:12px 28px;border-radius:6px;text-decoration:none;font-size:14px">
              Ir al sistema
            </a>
          </div>
        </div>
        <div style="padding:16px 24px;background:#f0f0f0;font-size:12px;color:#888;text-align:center">
          ENTREGAS INDUMENTARIA PEÑAROL
        </div>
      </div>
    `

    await transporter.sendMail({
      from: '"Indumentaria Peñarol" <edominguez@capenarol.com.uy>',
      to: 'edominguez@capenarol.com.uy',
      subject: `Nuevo usuario registrado: ${displayName}`,
      text: `Nuevo usuario registrado:\n\nNombre: ${displayName}\nEmail: ${email}\nTeléfono: ${telefono || '—'}\nCargo: ${cargo || '—'}${categoria ? `\nSector: ${categoria}${division ? ` / ${division}` : ''}` : ''}\n\nEl usuario ya puede ingresar al sistema.\n\nhttps://indumentaria-cap.vercel.app`,
      html,
    })

    res.status(200).json({ ok: true })

  } catch (err) {
    console.error('Error enviando notificación de registro:', err.message)
    res.status(500).json({ error: err.message })
  }
}
