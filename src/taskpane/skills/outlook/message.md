# Reading the Current Message — Mailbox read APIs

## Core Concept

Outlook uses the **Mailbox API** (`Office.context.mailbox`), not a host object model. There is **no `.run()` batch wrapper and no `context.sync()`** — that model only applies to Word, Excel, and PowerPoint.

- The current message is `Office.context.mailbox.item`.
- Item properties are either **plain properties** (read them directly) or **async objects** with `getAsync`.
- Every `...Async` call takes a **callback** `(result: Office.AsyncResult<T>) => void`. Check `result.status === Office.AsyncResultStatus.Succeeded`, then read `result.value`.

## Read the item identity and envelope

```js
const item = Office.context.mailbox.item;
const profile = Office.context.mailbox.userProfile;

return {
  itemType: item.itemType,                       // "message" | "appointment"
  subject: item.subject,                         // plain property
  normalizedSubject: item.normalizedSubject,     // reply/forward prefixes stripped
  from: item.from,                               // EmailAddressDetails: displayName, emailAddress
  sender: item.sender,
  to: item.to.map(r => r.emailAddress),          // Recipient[] -> EmailAddressDetails
  cc: item.cc.map(r => r.emailAddress),
  itemId: item.itemId,
  conversationId: item.conversationId,
  internetMessageId: item.internetMessageId,
  dateTimeCreated: item.dateTimeCreated,         // JS Date
  attachments: item.attachments.map(a => ({ name: a.name, size: a.size, contentType: a.contentType })),
  user: { name: profile.displayName, email: profile.emailAddress },
};
```

## Read the body (async, required)

The body is **always async** — in read mode it is read-only:

```js
const bodyHtml = await new Promise((resolve, reject) => {
  Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
    else reject(new Error(result.error && result.error.message));
  });
});
```

- `Office.CoercionType.Text` returns plain text (best for summarising, quoting, or keyword checks).
- `Office.CoercionType.Html` returns the raw HTML — strip tags with a regex before showing it to the user.

## Gotchas

1. **No `load()` / `context.sync()`** — if you write `item.load(...)` or `await context.sync()` the code fails. Use the property or the `getAsync` callback form.
2. **Read mode is read-only.** To modify the message use reply/forward forms — see the `compose` skill.
3. **Attachment contents are not available** with `ReadWriteItem`. Only metadata (name, size, contentType) is exposed; reading attachment bytes needs `ReadWriteMailbox` plus EWS/REST.
4. **`getAsync` may not run in a sandboxed iframe** — this add-in executes code in the task pane window, so `Office.context.mailbox` is available directly.
5. **Async callbacks are not awaited automatically** — always wrap them in a `Promise` as shown above, then `await` it.
6. **`item.body.getAsync` fails on items with no body** (rare) — guard with `if (result.value)` and return `''`.

## Return value tips

Return a plain, serialisable object (not the Office proxies). Round long bodies: `body.slice(0, 4000) + '…'` so the conversation stays small.
