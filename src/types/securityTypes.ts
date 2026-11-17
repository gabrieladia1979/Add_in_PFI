/**
 * Resultado de autenticación de email (SPF, DKIM, DMARC, compauth).
 */
export type AuthResult =
  | "pass"
  | "fail"
  | "softfail"
  | "neutral"
  | "none"
  | "bestguesspass"
  | "unknown";

/**
 * Features de seguridad extraídas de los headers MIME y propiedades
 * de Office.js para alimentar al clasificador de phishing.
 */
export interface SecurityFeatures {
  // === Autenticación (Authentication-Results header) ===
  spf_result: AuthResult;
  dkim_result: AuthResult;
  dmarc_result: AuthResult;
  compauth_result: AuthResult;

  // === Detección de Spoofing ===
  return_path: string;
  reply_to: string;
  from_return_path_match: boolean;
  from_reply_to_match: boolean;
  sender_vs_from_match: boolean;

  // === Anti-Spam de Microsoft 365 (X-Forefront-Antispam-Report) ===
  scl: number;
  threat_category: string;
  spam_filtering_verdict: string;
  auth_as: "Internal" | "Anonymous" | "unknown";
  bcl: number;

  // === Origen ===
  originating_ip: string;
  originating_country: string;
  x_mailer: string;
  received_hop_count: number;

  // === Adjuntos ===
  attachment_count: number;
  attachment_types: string[];
  has_executable_attachment: boolean;

  // === Destinatarios ===
  to_count: number;
  cc_count: number;

  // === Metadata ===
  internet_message_id: string;
  has_headers: boolean;
}
