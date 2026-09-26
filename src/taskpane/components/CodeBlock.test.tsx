import { describe, it, expect, afterEach } from 'vitest';
import { render, screen, cleanup } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { LanguageProvider } from '../i18n/index.ts';
import { CodeBlock } from './CodeBlock.tsx';

afterEach(cleanup);

const CODE = 'Office.context.mailbox.item.displayReplyFormAsync({ htmlBody: "<p>ok</p>" });';

function renderCodeBlock(ui: React.ReactElement) {
  return render(<LanguageProvider initialLocale="en">{ui}</LanguageProvider>);
}

describe('CodeBlock — approval summary', () => {
  it('shows the plain-language summary and hides the code behind a toggle', async () => {
    renderCodeBlock(
      <CodeBlock
        code={CODE}
        summary="Opens a reply with a suggested answer."
        status="pending"
        onApprove={() => {}}
        onReject={() => {}}
      />
    );

    expect(screen.getByText('Opens a reply with a suggested answer.')).toBeTruthy();
    expect(screen.queryByText(CODE)).toBeNull();

    await userEvent.click(await screen.findByRole('button', { name: 'Show code' }));

    expect(screen.getByText(CODE)).toBeTruthy();
    expect(screen.getByRole('button', { name: 'Hide code' })).toBeTruthy();
  });

  it('re-hides the code when the toggle is used again', async () => {
    renderCodeBlock(<CodeBlock code={CODE} summary="Drafts the reply." status="success" />);

    await userEvent.click(await screen.findByRole('button', { name: 'Show code' }));
    await userEvent.click(screen.getByRole('button', { name: 'Hide code' }));

    expect(screen.queryByText(CODE)).toBeNull();
  });

  it('renders without a summary', () => {
    renderCodeBlock(<CodeBlock code={'return 1;'} status="success" />);

    expect(screen.getByText(/return 1;/)).toBeTruthy();
    expect(screen.queryByRole('button', { name: 'Show code' })).toBeNull();
  });
});


