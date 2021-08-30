.env で`CLIENTID`と`SECRET`を設定する。

```
OAUTH_CLIENT_ID=CLIENTID
OAUTH_CLIENT_SECRET=SECRET
OAUTH_REDIRECT_URI=http://localhost:3000/auth/callback
OAUTH_SCOPES='user.read,calendars.readwrite,mailboxsettings.read,presence.read,presence.read.all'
OAUTH_AUTHORITY=https://login.microsoftonline.com/common/
```