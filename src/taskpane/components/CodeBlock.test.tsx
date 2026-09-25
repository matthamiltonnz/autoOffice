import { describe, it, expect, afterEach } from 'vitest';
import { render, screen, cleanup } from '@testing-library/react';
import { LanguageProvider } from '../i18n/index.ts';
import { CodeBlock } from './CodeBlock.tsx';

afterEach(cleanup);

function renderCodeBlock(ui: React.ReactElement) {
  return render(<LanguageProvider initialLocale="en">{ui}</LanguageProvider>);
}

describe('CodeBlock — approval summary', () => {
  it('shows the plain-language summary above the code', () => {
    renderCodeBlock(
      <CodeBlock
        code={'Office.context.mailbox.item.displayReplyFormAsync({ htmlBody: "<p>ok</p>" });'}
        summary="Opens a reply with a suggested answer."
        status="pending"
        onApprove={() => {}}
        onReject={() => {}}
      />
    );

    expect(screen.getByText('Opens a reply with a suggested answer.')).toBeTruthy();
    expect(screen.getByText(/displayReplyFormAsync/)).toBeTruthy();
  });

  it('renders without a summary', () => {
    renderCodeBlock(<CodeBlock code={'return 1;'} status="success" />);

    expect(screen.getByText(/return 1;/)).toBeTruthy();
  });
});

