# Recipients, Appointments, and Item Metadata

## Recipients (message items)

Recipient fields are `Recipient[]`; each entry has `displayName`, `emailAddress`, and for compose items a `recipientType`.

```js
const item = Office.context.mailbox.item;

// Read (works in read and compose mode)
const to  = item.to.map(r => r.emailAddress);
const cc  = item.cc.map(r => r.emailAddress);
const bcc = item.bcc.map(r => r.emailAddress);   // compose mode only for bcc writes

// Write (compose mode only)
item.to.setAsync(['ops@positiveit.co.nz'], (result) => {
  // result.status === Office.AsyncResultStatus.Succeeded
});
item.cc.setAsync(['accounts@positiveit.co.nz'], () => {});
```

`item.to`/`cc`/`bcc` are also async objects — the `getAsync` form works on older clients:

```js
item.to.getAsync((result) => {
  const list = result.value.map(r => r.emailAddress);
});
```

## Appointments (meeting items)

```js
const item = Office.context.mailbox.item;
const start = await new Promise((res) => item.start.getAsync((r) => res(r.value)));  // Date
const end   = await new Promise((res) => item.end.getAsync((r) => res(r.value)));
item.location;                     // string
item.requiredAttendees.map(a => a.emailAddress);
item.optionalAttendees.map(a => a.emailAddress);
item.organizer;                    // EmailAddressDetails
```

Compose mode adds `setAsync` for `start`, `end`, `location`, `requiredAttendees`, `optionalAttendees`. Note appointment properties are exposed only to the **organizer's** side; attendees see a read-only view.

## Item metadata worth returning

| Property | Meaning |
| --- | --- |
| `item.itemType` | `"message"` or `"appointment"` |
| `item.itemClass` | `IPM.Note`, `IPM.Schedule.Meeting.Request`, … |
| `item.conversationId` | groups the thread (use it to match a conversation) |
| `item.internetMessageId` | stable RFC-5322 id, good for matching an exported thread |
| `item.dateTimeCreated` / `dateTimeModified` | JS `Date` |
| `item.itemId` | changes when the item is moved between folders |
| `Office.context.mailbox.diagnostics.hostName` | `Outlook` |

## Gotchas

1. **Read mode allows reads only** — `setAsync` on `to`/`cc` needs a compose (reply/forward/new) item.
2. **`bcc` is not returned in read mode** for privacy; treat it as compose-only.
3. **Recipient arrays can be empty** (`[]`), never `undefined` on modern clients — still guard with `(item.to || [])`.
4. **`itemId` is unstable.** Persisting it across folders/sessions breaks; persist `internetMessageId` or `conversationId` instead.
5. **Meeting invites**: for a calendar item, use `item.start`/`item.end` (`Date` objects) rather than parsing the body.
6. **Mailbox-wide search is out of scope** — this add-in is granted `ReadWriteItem` (current item only). Reading other messages, folders, or search results needs the `ReadWriteMailbox` permission plus EWS/REST calls via `getCallbackTokenAsync`.
