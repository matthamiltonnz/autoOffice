# Writing to the Message — compose and reply APIs

## Which mode am I in?

```js
const item = Office.context.mailbox.item;
// Read mode  (a received message is open):
//   item.body is read-only, use displayReplyFormAsync / displayNewMessageFormAsync.
// Compose mode (a reply/forward/new message is being written):
//   item.body.setAsync, item.subject.setAsync, item.to.setAsync are available.
```

Detect it defensively by calling `item.body.getTypeAsync` and checking what exists:

```js
const canWrite = typeof item.body.setAsync === 'function';
```

## Compose mode — set the body and subject

```js
function setBody(html) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.setAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      (result) => result.status === Office.AsyncResultStatus.Succeeded
        ? resolve('body updated')
        : reject(new Error(result.error && result.error.message))
    );
  });
}

function setSubject(text) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.subject.setAsync(text, (result) =>
      result.status === Office.AsyncResultStatus.Succeeded
        ? resolve('subject updated')
        : reject(new Error(result.error && result.error.message)));
  });
}

const output = await setBody('<p>Thanks — confirmed for Tuesday.</p>');
```

Useful body helpers: `prependAsync(content, options, cb)`, `appendAsync(...)`, `setSelectedDataAsync(...)`, and `getTypeAsync(cb)` to learn whether the draft is HTML or plain text.

## Read mode — open a reply/forward prefilled

```js
Office.context.mailbox.item.displayReplyFormAsync(
  { htmlBody: '<p>Hi — yes, Tuesday works. Full answer below.</p>' },
  (result) => { /* result.status */ }
);
```

Variants: `displayReplyAllFormAsync`, and `displayForwardFormAsync` (compose a forward), plus `displayNewMessageFormAsync({ to, cc, subject, htmlBody })` (Mailbox 1.1+) to start a brand-new message.

## Attachments and saving

```js
// Attach a file by URL (must be reachable by Outlook).
Office.context.mailbox.item.addFileAttachmentAsync(uri, 'report.pdf', (result) => { /* result.value = attachment id */ });
// Attach another message/meeting by itemId (e.g. one you read earlier).
Office.context.mailbox.item.addItemAttachmentAsync(itemId, 'original.msg', (result) => {});
// Save and close a draft:
Office.context.mailbox.item.saveAsync((result) => { /* persisted, result.value = itemId */ });
Office.context.mailbox.item.closeAsync();
```

## Gotchas

1. **Do not send the message.** Sending is left to the user — `saveAsync`/`closeAsync` only persist or dismiss the draft.
2. **`setAsync` replaces the whole body.** Read it first (`item.body.getAsync`) when you want to append to the user's existing text.
3. **Coercion type must match content.** Passing plain text with `CoercionType.Html` shows `<p>`-less text but breaks newlines; prefer `Html` and wrap paragraphs in `<p>`.
4. **`addFileAttachmentAsync` needs a URL the Exchange server can fetch**, not a `blob:` or `data:` URL.
5. **Read-mode write attempts fail with `InvalidAccessError`** — switch to a reply form instead.
6. **Await the callbacks** (wrap in a Promise) or the task pane will report success before the write lands.

## Compose mode: writing into the draft that is already open

If the user is composing or replying, the item is writable and no reply-form
API is involved — `displayReplyFormAsync` does **not** exist here. Detect the
mode from the capabilities line in your instructions, never by probing.

```js
const item = Office.context.mailbox.item;
const html = '<p>Hello,</p><p>Thank you for your message.</p>';

await new Promise((resolve, reject) => {
  item.body.setAsync(html, { coercionType: Office.CoercionType.Html }, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
    else reject(new Error(result.error?.message ?? 'body.setAsync failed'));
  });
});

// subject.setAsync(...) only if the user asked for a subject.
return 'Replaced the body of the open draft — review it and hit Send.';
```

Add recipients with `item.to.addAsync(...)` / `cc.addAsync(...)`, never
`displayReplyFormAsync`. If an execution errors, report it plainly — do not
claim the draft was written.

