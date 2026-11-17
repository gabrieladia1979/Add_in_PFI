export type MailMetadata = {
  sender: string;
  senderEmail: string;
  subject: string;
  body: string;
  bodyPreview: string;
};

const BODY_PREVIEW_LIMIT = 280;

export function getMailboxItem(): Office.MessageRead | null {
  if (typeof Office === "undefined") {
    return null;
  }

  const item = Office.context?.mailbox?.item;
  if (!item) {
    return null;
  }

  return item as Office.MessageRead;
}

function getSenderLabel(item: Office.MessageRead): string {
  const from = item.from;
  if (from?.displayName) {
    return from.displayName;
  }
  if (from?.emailAddress) {
    return from.emailAddress;
  }

  return "Desconocido";
}

function normalizeBodyPreview(text: string): string {
  const compact = text.replace(/\s+/g, " ").trim();
  if (!compact) {
    return "Sin contenido";
  }

  return compact.slice(0, BODY_PREVIEW_LIMIT);
}

function getBodyText(item: Office.MessageRead): Promise<string> {
  if (!item.body) {
    return Promise.resolve("");
  }

  return new Promise((resolve) => {
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || "");
        return;
      }

      resolve("");
    });
  });
}

function getBodyHtml(item: Office.MessageRead): Promise<string> {
  if (!item.body) {
    return Promise.resolve("");
  }

  return new Promise((resolve) => {
    item.body.getAsync(Office.CoercionType.Html, (result) => {
      resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value || "" : "");
    });
  });
}

export async function getCurrentMessageMetadata(): Promise<MailMetadata | null> {
  const item = getMailboxItem();
  if (!item) {
    return null;
  }

  const sender = getSenderLabel(item);
  const senderEmail = item.from?.emailAddress || "";
  const subject = item.subject || "Sin asunto";
  const [bodyText, bodyHtml] = await Promise.all([getBodyText(item), getBodyHtml(item)]);

  return {
    sender,
    senderEmail,
    subject,
    // Preserve href targets: the visible link text alone may hide a phishing URL.
    body: bodyHtml || bodyText,
    bodyPreview: normalizeBodyPreview(bodyText),
  };
}

function escapeHtml(value: string): string {
  return value.replace(/[&<>"']/g, (character) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;",
  })[character] || character);
}

/** Abre un borrador para el área de seguridad. El usuario revisa y envía el reporte. */
export function prepareIncidentReport(recipient: string, metadata: MailMetadata, riskScore?: number): void {
  const item = getMailboxItem();
  if (!item?.itemId) {
    throw new Error("No se pudo adjuntar el correo original. Abrí un mensaje y volvé a intentar.");
  }

  const summary = riskScore === undefined
    ? "Reporte manual de un correo sospechoso."
    : `Resultado de PhishARG: riesgo estimado ${Math.round(riskScore)} %. Revisar el correo original adjunto.`;

  Office.context.mailbox.displayNewMessageForm({
    toRecipients: [recipient],
    subject: `Reporte de posible phishing: ${metadata.subject}`.slice(0, 255),
    htmlBody: `<p>Solicito revisar este posible incidente de phishing.</p><p>${escapeHtml(summary)}</p><p>Remitente: ${escapeHtml(metadata.senderEmail || metadata.sender)}<br>Asunto: ${escapeHtml(metadata.subject)}</p><p>El correo original está adjunto. Revisá el destinatario antes de enviarlo.</p>`,
    attachments: [{
      type: Office.MailboxEnums.AttachmentType.Item,
      name: "Correo sospechoso",
      itemId: item.itemId,
    }],
  });
}
