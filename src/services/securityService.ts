import { SecurityFeatures, AuthResult } from "../types/securityTypes";

/** Extensiones de archivo consideradas ejecutables peligrosas */
const DANGEROUS_EXTENSIONS = new Set([
  ".exe", ".bat", ".cmd", ".com", ".vbs", ".vbe",
  ".js", ".jse", ".wsf", ".wsh", ".msi", ".scr",
  ".pif", ".ps1", ".hta", ".cpl", ".reg", ".inf",
]);

// ─── Header Retrieval ──────────────────────────────────────────────

/**
 * Obtiene todos los headers MIME del email actual.
 * Requiere Mailbox requirement set 1.8.
 */
function getInternetHeaders(item: Office.MessageRead): Promise<string> {
  return new Promise((resolve) => {
    let settled = false;
    const complete = (val: string): void => {
      if (!settled) {
        settled = true;
        window.clearTimeout(timerId);
        resolve(val);
      }
    };
    const timerId = window.setTimeout(() => complete(""), 4000);

    if (
      !Office.context.requirements.isSetSupported("Mailbox", "1.8") ||
      !item.getAllInternetHeadersAsync
    ) {
      complete("");
      return;
    }

    try {
      item.getAllInternetHeadersAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          complete(result.value || "");
          return;
        }
        complete("");
      });
    } catch {
      complete("");
    }
  });
}

// ─── Parsing Helpers ───────────────────────────────────────────────

/**
 * Extrae el dominio de una dirección de email.
 * "user@example.com" -> "example.com"
 */
function extractDomain(email: string): string {
  const atIndex = email.lastIndexOf("@");
  if (atIndex === -1) {
    return "";
  }
  return email.substring(atIndex + 1).toLowerCase().trim();
}

/**
 * Extrae una dirección de email de un string con formato variado.
 * Maneja: "<user@example.com>", "User Name <user@example.com>", "user@example.com"
 */
function extractEmailFromHeader(value: string): string {
  const angleMatch = value.match(/<([^>]+)>/);
  if (angleMatch) {
    return angleMatch[1].trim().toLowerCase();
  }
  const emailMatch = value.match(/[\w.+-]+@[\w.-]+\.[a-zA-Z]{2,}/);
  if (emailMatch) {
    return emailMatch[0].toLowerCase();
  }
  return value.trim().toLowerCase();
}

/**
 * Parsea los resultados de autenticación del header Authentication-Results.
 */
function parseAuthenticationResults(headers: string): {
  spf: AuthResult;
  dkim: AuthResult;
  dmarc: AuthResult;
  compauth: AuthResult;
} {
  const result = {
    spf: "unknown" as AuthResult,
    dkim: "unknown" as AuthResult,
    dmarc: "unknown" as AuthResult,
    compauth: "unknown" as AuthResult,
  };

  const spfMatch = headers.match(/spf=(pass|fail|softfail|neutral|none)/i);
  if (spfMatch) {
    result.spf = spfMatch[1].toLowerCase() as AuthResult;
  }

  const dkimMatch = headers.match(/dkim=(pass|fail|neutral|none)/i);
  if (dkimMatch) {
    result.dkim = dkimMatch[1].toLowerCase() as AuthResult;
  }

  const dmarcMatch = headers.match(
    /dmarc=(pass|fail|bestguesspass|none)/i
  );
  if (dmarcMatch) {
    result.dmarc = dmarcMatch[1].toLowerCase() as AuthResult;
  }

  const compMatch = headers.match(/compauth=(pass|fail)/i);
  if (compMatch) {
    result.compauth = compMatch[1].toLowerCase() as AuthResult;
  }

  return result;
}

/**
 * Parsea los headers anti-spam de Microsoft 365.
 * Extrae campos del X-Forefront-Antispam-Report y X-Microsoft-Antispam.
 */
function parseAntispamHeaders(headers: string): {
  scl: number;
  threatCategory: string;
  sfv: string;
  bcl: number;
  originatingCountry: string;
  connectingIp: string;
} {
  const result = {
    scl: -1,
    threatCategory: "NONE",
    sfv: "",
    bcl: -1,
    originatingCountry: "",
    connectingIp: "",
  };

  // SCL from X-Forefront-Antispam-Report
  const sclMatch = headers.match(/SCL:(\-?\d+)/i);
  if (sclMatch) {
    result.scl = parseInt(sclMatch[1], 10);
  }

  // CAT (Threat Category)
  const catMatch = headers.match(/CAT:([A-Z]+)/i);
  if (catMatch) {
    result.threatCategory = catMatch[1].toUpperCase();
  }

  // SFV (Spam Filtering Verdict)
  const sfvMatch = headers.match(/SFV:([A-Z]+)/i);
  if (sfvMatch) {
    result.sfv = sfvMatch[1].toUpperCase();
  }

  // BCL from X-Microsoft-Antispam
  const bclMatch = headers.match(/BCL:(\d+)/i);
  if (bclMatch) {
    result.bcl = parseInt(bclMatch[1], 10);
  }

  // CTRY (Country)
  const ctryMatch = headers.match(/CTRY:([A-Z]{2})/i);
  if (ctryMatch) {
    result.originatingCountry = ctryMatch[1].toUpperCase();
  }

  // CIP (Connecting IP)
  const cipMatch = headers.match(/CIP:([\d.]+)/i);
  if (cipMatch) {
    result.connectingIp = cipMatch[1];
  }

  return result;
}

/**
 * Parsea indicadores de spoofing comparando From con Return-Path y Reply-To.
 */
function parseSpoofingIndicators(
  headers: string,
  fromEmail: string
): {
  returnPath: string;
  replyTo: string;
  fromReturnPathMatch: boolean;
  fromReplyToMatch: boolean;
} {
  const fromDomain = extractDomain(fromEmail);

  // Return-Path
  const returnPathMatch = headers.match(/Return-Path:\s*<?([^>\r\n]+)>?/im);
  const returnPath = returnPathMatch
    ? extractEmailFromHeader(returnPathMatch[1])
    : "";
  const returnPathDomain = extractDomain(returnPath);

  // Reply-To
  const replyToMatch = headers.match(/Reply-To:\s*<?([^>\r\n]+)>?/im);
  const replyTo = replyToMatch
    ? extractEmailFromHeader(replyToMatch[1])
    : "";
  const replyToDomain = extractDomain(replyTo);

  return {
    returnPath,
    replyTo,
    fromReturnPathMatch:
      returnPath === "" || fromDomain === returnPathDomain,
    fromReplyToMatch:
      replyTo === "" || fromDomain === replyToDomain,
  };
}

/**
 * Parsea metadata de origen: IP, X-Mailer, cantidad de hops.
 */
function parseOriginMetadata(headers: string): {
  originatingIp: string;
  xMailer: string;
  receivedHopCount: number;
} {
  // X-Originating-IP
  const origIpMatch = headers.match(
    /X-Originating-IP:\s*\[?([^\]\r\n]+)\]?/im
  );
  const originatingIp = origIpMatch ? origIpMatch[1].trim() : "";

  // X-Mailer or User-Agent
  const mailerMatch = headers.match(/(?:X-Mailer|User-Agent):\s*([^\r\n]+)/im);
  const xMailer = mailerMatch ? mailerMatch[1].trim() : "";

  // Count Received: hops
  const receivedMatches = headers.match(/^Received:/gim);
  const receivedHopCount = receivedMatches ? receivedMatches.length : 0;

  return { originatingIp, xMailer, receivedHopCount };
}

/**
 * Extrae el tipo de autenticación de la organización.
 */
function parseAuthAs(headers: string): "Internal" | "Anonymous" | "unknown" {
  const match = headers.match(
    /X-MS-Exchange-Organization-AuthAs:\s*(Internal|Anonymous)/im
  );
  if (match) {
    return match[1] as "Internal" | "Anonymous";
  }
  return "unknown";
}

// ─── Attachment Analysis ───────────────────────────────────────────

/**
 * Analiza los adjuntos del email para extraer features de seguridad.
 */
function getAttachmentFeatures(item: Office.MessageRead): {
  count: number;
  types: string[];
  hasExecutable: boolean;
} {
  const attachments = item.attachments || [];
  const types: string[] = [];
  let hasExecutable = false;

  for (const att of attachments) {
    if (att.contentType) {
      types.push(att.contentType);
    }

    // Verificar extensión peligrosa en el nombre del adjunto
    const name = (att.name || "").toLowerCase();
    const dotIndex = name.lastIndexOf(".");
    if (dotIndex !== -1) {
      const ext = name.substring(dotIndex);
      if (DANGEROUS_EXTENSIONS.has(ext)) {
        hasExecutable = true;
      }
    }
  }

  return {
    count: attachments.length,
    types,
    hasExecutable,
  };
}

// ─── Main Extraction ───────────────────────────────────────────────

/**
 * Crea un objeto SecurityFeatures con valores por defecto (sin datos).
 */
function createEmptyFeatures(): SecurityFeatures {
  return {
    spf_result: "unknown",
    dkim_result: "unknown",
    dmarc_result: "unknown",
    compauth_result: "unknown",
    return_path: "",
    reply_to: "",
    from_return_path_match: true,
    from_reply_to_match: true,
    sender_vs_from_match: true,
    scl: -1,
    threat_category: "NONE",
    spam_filtering_verdict: "",
    auth_as: "unknown",
    bcl: -1,
    originating_ip: "",
    originating_country: "",
    x_mailer: "",
    received_hop_count: 0,
    attachment_count: 0,
    attachment_types: [],
    has_executable_attachment: false,
    to_count: 0,
    cc_count: 0,
    internet_message_id: "",
    has_headers: false,
  };
}

/**
 * Extrae todas las features de seguridad del email actualmente abierto.
 * Esta es la función principal que orquesta todo el proceso de extracción.
 *
 * @param item - El item de Office.js MessageRead
 * @returns SecurityFeatures con todos los datos extraídos
 */
export async function extractSecurityFeatures(
  item: Office.MessageRead
): Promise<SecurityFeatures> {
  const features = createEmptyFeatures();

  // --- Propiedades directas de Office.js ---

  // Sender vs From
  const fromEmail = item.from?.emailAddress?.toLowerCase() || "";
  const senderEmail = item.sender?.emailAddress?.toLowerCase() || "";
  features.sender_vs_from_match =
    fromEmail === senderEmail || senderEmail === "";

  // Internet Message ID
  features.internet_message_id = item.internetMessageId || "";

  // Adjuntos
  const attachInfo = getAttachmentFeatures(item);
  features.attachment_count = attachInfo.count;
  features.attachment_types = attachInfo.types;
  features.has_executable_attachment = attachInfo.hasExecutable;

  // Destinatarios
  features.to_count = (item.to || []).length;
  features.cc_count = (item.cc || []).length;

  // --- Headers MIME ---
  const headers = await getInternetHeaders(item);

  if (!headers) {
    features.has_headers = false;
    return features;
  }

  features.has_headers = true;

  // Autenticación
  const auth = parseAuthenticationResults(headers);
  features.spf_result = auth.spf;
  features.dkim_result = auth.dkim;
  features.dmarc_result = auth.dmarc;
  features.compauth_result = auth.compauth;

  // Anti-spam
  const antispam = parseAntispamHeaders(headers);
  features.scl = antispam.scl;
  features.threat_category = antispam.threatCategory;
  features.spam_filtering_verdict = antispam.sfv;
  features.bcl = antispam.bcl;
  features.originating_country = antispam.originatingCountry;

  // AuthAs
  features.auth_as = parseAuthAs(headers);

  // Spoofing
  const spoofing = parseSpoofingIndicators(headers, fromEmail);
  features.return_path = spoofing.returnPath;
  features.reply_to = spoofing.replyTo;
  features.from_return_path_match = spoofing.fromReturnPathMatch;
  features.from_reply_to_match = spoofing.fromReplyToMatch;

  // Origen
  const origin = parseOriginMetadata(headers);
  features.originating_ip =
    origin.originatingIp || antispam.connectingIp;
  features.x_mailer = origin.xMailer;
  features.received_hop_count = origin.receivedHopCount;

  return features;
}
