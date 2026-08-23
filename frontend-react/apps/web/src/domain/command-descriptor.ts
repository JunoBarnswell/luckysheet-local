/**
 * UI-facing description of one command dispatch.
 *
 * The descriptor is the command contract itself: callers provide the runtime
 * command id and its parameters directly. There is intentionally no action
 * vocabulary or translation table in the web application.
 */
export interface CommandDescriptor<Params = unknown> {
  readonly commandId: string;
  readonly params?: Params;
}
