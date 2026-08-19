/** Structured application errors with user-safe messages and stable codes. */
export type AppErrorCode =
  | "UNSUPPORTED_FILE_TYPE"
  | "FILE_TOO_LARGE"
  | "CORRUPT_DOCUMENT"
  | "PASSWORD_PROTECTED"
  | "MACRO_DOCUMENT_REJECTED"
  | "UNSAFE_PACKAGE"
  | "DOCUMENT_NOT_FOUND"
  | "SESSION_EXPIRED"
  | "INVALID_REQUEST"
  | "PROCESSING_FAILED"
  | "VERIFICATION_UNAVAILABLE"
  | "NOT_READY"
  | "INTERNAL";

export class AppError extends Error {
  readonly code: AppErrorCode;
  readonly httpStatus: number;
  /** Message safe to show to end users (never a stack trace). */
  readonly userMessage: string;

  constructor(code: AppErrorCode, userMessage: string, httpStatus = 400) {
    super(`${code}: ${userMessage}`);
    this.code = code;
    this.userMessage = userMessage;
    this.httpStatus = httpStatus;
  }
}

export const Errors = {
  unsupportedType: () =>
    new AppError(
      "UNSUPPORTED_FILE_TYPE",
      "This version currently supports Microsoft Word .docx files.",
      415
    ),
  tooLarge: (maxBytes: number) =>
    new AppError(
      "FILE_TOO_LARGE",
      `File exceeds the maximum size of ${Math.round(maxBytes / (1024 * 1024))} MB.`,
      413
    ),
  corrupt: () =>
    new AppError(
      "CORRUPT_DOCUMENT",
      "The document appears to be corrupt or is not a valid Word .docx file.",
      422
    ),
  passwordProtected: () =>
    new AppError(
      "PASSWORD_PROTECTED",
      "The document is password protected. Please remove the password and try again.",
      422
    ),
  macroRejected: () =>
    new AppError(
      "MACRO_DOCUMENT_REJECTED",
      "Macro-enabled documents are not supported. Please save the file as a standard .docx.",
      422
    ),
  unsafePackage: (reason: string) =>
    new AppError(
      "UNSAFE_PACKAGE",
      `The document could not be safely processed (${reason}).`,
      422
    ),
  notFound: () =>
    new AppError(
      "DOCUMENT_NOT_FOUND",
      "Document not found. It may have been deleted by the retention policy.",
      404
    ),
  invalid: (msg: string) => new AppError("INVALID_REQUEST", msg, 400),
  notReady: (msg: string) => new AppError("NOT_READY", msg, 409),
  internal: () =>
    new AppError(
      "INTERNAL",
      "Something went wrong while processing the document.",
      500
    ),
};
