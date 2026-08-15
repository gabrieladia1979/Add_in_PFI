export type MailMetadata = {
  sender: string;
  subject: string;
  body: string;
  bodyPreview: string;
};

const BODY_PREVIEW_LIMIT = 280;

function getMailboxItem(): Office.MessageRead | null {
  if (typeof Office === "undefined") {
    return null;
  }

  const item = Office.context?.mailbox?.item;
  if (!item) {
    return null;
  }

  return item as Office.MessageRead;
}

function getSenderEmail(item: Office.MessageRead): string {
  return item.from?.emailAddress || "Desconocido";
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

  return new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || "");
        return;
      }

      reject(new Error(result.error?.message || "No se pudo leer el cuerpo del correo"));
    });
  });
}

export async function getCurrentMessageMetadata(): Promise<MailMetadata | null> {
  const item = getMailboxItem();
  if (!item) {
    return null;
  }

  const sender = getSenderEmail(item);
  const subject = item.subject || "Sin asunto";
  const bodyText = await getBodyText(item);

  return {
    sender,
    subject,
    body: bodyText,
    bodyPreview: normalizeBodyPreview(bodyText),
  };
}
