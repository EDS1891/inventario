import nodemailer from 'nodemailer'

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*')
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS')
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type')
  if (req.method === 'OPTIONS') return res.status(200).end()

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' })
  }

  const { type, to, creadorDisplayName, receptorNombre, delId, lines, persona, motivo } = req.body

  if (!to || !type || !delId) {
    return res.status(400).json({ error: 'Faltan datos' })
  }

  const esAceptado = type === 'aceptado'
  const accionLabel = esAceptado ? 'ACEPTADA' : 'RECHAZADA'
  const accionColor = esAceptado ? '#2e9b5e' : '#C2473D'
  const accionColorBg = esAceptado ? '#edf7f2' : '#fdf0ef'
  const accionColorBorder = esAceptado ? '#a8dfc2' : '#f0b8b4'
  const accionEmoji = esAceptado ? '✓' : '✕'

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

    const linesHtml = (lines || []).map(l => `
      <div style="display:flex;justify-content:space-between;padding:6px 0;border-bottom:1px solid #eee;font-size:14px">
        <span>${l.name} – Talle ${l.talle}</span>
        <span style="font-weight:bold">x${l.qty}</span>
      </div>`).join('')

    const motivoHtml = !esAceptado && motivo ? `
      <div style="margin-top:16px;padding:12px 16px;background:#fff4f4;border:1px solid #f0b8b4;border-radius:8px;font-size:13px;color:#444">
        <strong>Motivo del rechazo:</strong><br/>${motivo}
      </div>` : ''

    const html = `
      <div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto">
        <div style="background:#121212;padding:24px;text-align:center">
          <div style="color:#FFD200;font-weight:bold;font-size:18px;letter-spacing:1px">INDUMENTARIA PEÑAROL</div>
        </div>
        <div style="padding:28px 24px;background:#fff;border:1px solid #e0e0e0">
          <p style="margin:0 0 12px;font-size:15px">Hola <strong>${creadorDisplayName || to}</strong>,</p>
          <div style="background:${accionColorBg};border:1px solid ${accionColorBorder};border-radius:8px;padding:14px 18px;margin-bottom:20px;font-size:15px;font-weight:bold;color:${accionColor}">
            ${accionEmoji} La entrega #${delId} fue <strong>${accionLabel}</strong> por ${receptorNombre || 'el receptor'}.
          </div>
          <div style="background:#f8f8f8;border:1px solid #e0e0e0;border-radius:8px;padding:16px;margin-bottom:16px">
            <div style="font-weight:bold;font-size:12px;color:#888;letter-spacing:.05em;margin-bottom:6px">ENTREGA #${delId} — ${persona || ''}</div>
            ${linesHtml}
          </div>
          ${motivoHtml}
          <div style="margin-top:20px">
            <a href="https://indumentaria-cap.vercel.app" style="display:inline-block;background:#FFD200;color:#121212;font-weight:bold;padding:12px 28px;border-radius:6px;text-decoration:none;font-size:14px">
              Ver en el sistema
            </a>
          </div>
        </div>
        <div style="padding:16px 24px;background:#f0f0f0;font-size:12px;color:#888;text-align:center">
          ENTREGAS INDUMENTARIA PEÑAROL
        </div>
      </div>
    `

    const textMotivo = !esAceptado && motivo ? `\nMotivo: ${motivo}` : ''
    const linesTxt = (lines || []).map(l => `• ${l.name} – Talle ${l.talle} (x${l.qty})`).join('\n')

    await transporter.sendMail({
      from: '"Indumentaria Peñarol" <edominguez@capenarol.com.uy>',
      to,
      subject: `Entrega #${delId} ${accionLabel} — ${receptorNombre || 'Receptor'}`,
      text: `Hola ${creadorDisplayName || to},\n\nLa entrega #${delId} fue ${accionLabel} por ${receptorNombre || 'el receptor'}.${textMotivo}\n\nDetalle:\n${linesTxt}\n\nIngresá a https://indumentaria-cap.vercel.app\n\nDepósito Peñarol`,
      html,
    })

    res.status(200).json({ ok: true })

  } catch (err) {
    console.error('Error enviando notificación de confirmación:', err.message)
    res.status(500).json({ error: err.message })
  }
}
