export type MailMetadata = {
  sender: string;
  subject: string;
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

export async function getCurrentMessageMetadata(): Promise<MailMetadata | null> {
  const item = getMailboxItem();
  if (!item) {
    return null;
  }

  const sender = getSenderLabel(item);
  const subject = item.subject || "Sin asunto";
  const bodyText = await getBodyText(item);

  return {
    sender,
    subject,
    bodyPreview: normalizeBodyPreview(bodyText),
  };
}
