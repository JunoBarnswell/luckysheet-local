export interface LocalUser {
  readonly id: string;
  readonly username: string;
  readonly displayName: string;
  readonly enabled: boolean;
  readonly admin: boolean;
}
export interface IdentityActions {
  listUsers(): Promise<readonly LocalUser[]>;
  createUser(input: { username: string; displayName: string; password: string }): Promise<void>;
  setUserEnabled(userId: string, enabled: boolean): Promise<void>;
  resetPassword(userId: string, password: string): Promise<void>;
}
