# Security Policy

## Reporting a Vulnerability

Please **do not** open a public GitHub issue for security vulnerabilities.

See our security policy and contact details at:
**https://retyc.com/.well-known/security.txt**

## Known limitations

- Authentication tokens (access token + refresh token) are stored in
  `Office.context.roamingSettings`, which is mailbox-scoped and synchronized by
  Microsoft 365 across the user's devices. Anyone with access to the mailbox
  (account credentials or admin delegation) can read them. Protect your account
  with strong authentication and MFA.
