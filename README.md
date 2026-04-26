# EES WO Control Database — Day 23

Replace your current files with this package, deploy `firebase.rules.json`, then redeploy Netlify.

Important: login/session history starts from the moment this version is deployed. Old logins before Day 23 will not exist in `ees_sessions`.

Flux now has operational read access to:
- work orders
- audit logs
- presence
- leaves
- sessions
- attendance

Flux approved write actions include:
- WO priority updates
- task status/progress/remarks
- assign/remove assignees
- attendance status updates

Gemini key remains only in Netlify environment variables as `GEMINI_API_KEY`.
