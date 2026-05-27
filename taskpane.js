const SETTINGS_KEY = 'copilotDirectFlowAddin.settings';
let currentMail = null;

function $(id) { return document.getElementById(id); }

function defaultSettings() {
  return {
    agentName: 'MailTriageAgent',
    mode: 'summarize',
    flowUrl: '',
    authHeaderName: '',
    authHeaderValue: '',
    includeBody: false,
  };
}

function getSettings() {
  const raw = localStorage.getItem(SETTINGS_KEY);
  if (!raw) return defaultSettings();
  try {
    return { ...defaultSettings(), ...JSON.parse(raw) };
  } catch {
    return defaultSettings();
  }
}

function applySettings() {
  const s = getSettings();
  $('agentName').value = s.agentName;
  $('mode').value = s.mode;
  $('flowUrl').value = s.flowUrl;
  $('authHeaderName').value = s.authHeaderName;
  $('authHeaderValue').value = s.authHeaderValue;
  $('includeBody').checked = Boolean(s.includeBody);
}

function saveSettings() {
  const s = {
    agentName: $('agentName').value.trim() || 'MailTriageAgent',
    mode: $('mode').value,
    flowUrl: $('flowUrl').value.trim(),
    authHeaderName: $('authHeaderName').value.trim(),
    authHeaderValue: $('authHeaderValue').value.trim(),
    includeBody: $('includeBody').checked,
  };
  localStorage.setItem(SETTINGS_KEY, JSON.stringify(s));
  renderRequestPreview();
  setResult('設定を保存しました。');
}

function setResult(text) { $('result').textContent = text; }

function truncate(text, size = 120) {
  if (!text) return '-';
  return text.length <= size ? text : `${text.slice(0, size)}...`;
}

async function getBodyIfNeeded(includeBody) {
  if (!includeBody) return null;
  const item = Office.context.mailbox.item;
  return new Promise((resolve) => {
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value.slice(0, 8000));
      } else {
        resolve(null);
      }
    });
  });
}

async function readCurrentMail() {
  const item = Office.context.mailbox.item;
  if (!item) {
    $('mailStatus').textContent = '未選択';
    setResult('対象メールを開いた状態で実行してください。');
    return;
  }

  const rawItemId = item.itemId || null;
  let restId = null;
  try {
    if (rawItemId && Office.context.mailbox.convertToRestId) {
      restId = Office.context.mailbox.convertToRestId(rawItemId, Office.MailboxEnums.RestVersion.v2_0);
    }
  } catch (error) {
    console.warn('convertToRestId failed', error);
  }

  const from = item.from?.emailAddress || item.from?.displayName || '-';
  currentMail = {
    itemId: rawItemId,
    restId,
    subject: item.subject || '',
    from,
  };

  $('subject').textContent = truncate(currentMail.subject, 500);
  $('from').textContent = currentMail.from || '-';
  $('itemId').textContent = currentMail.itemId || '-';
  $('restId').textContent = currentMail.restId || '-';
  $('mailStatus').textContent = currentMail.itemId ? '取得済み' : 'ID未取得';
  renderRequestPreview();
}

function buildPayload(extra = {}) {
  const settings = getSettings();
  return {
    agentName: settings.agentName,
    mode: settings.mode,
    mailboxUser: Office.context.mailbox.userProfile?.emailAddress || null,
    message: {
      itemId: currentMail?.itemId || null,
      restId: currentMail?.restId || null,
      subject: currentMail?.subject || null,
      from: currentMail?.from || null,
      bodyIncluded: Boolean(settings.includeBody),
      body: extra.body ?? null,
    },
  };
}

function renderRequestPreview(extra = {}) {
  $('requestPreview').textContent = JSON.stringify(buildPayload(extra), null, 2);
}

async function runFlow() {
  const settings = getSettings();
  if (!settings.flowUrl) {
    setResult('Flow URL を設定してください。');
    return;
  }
  if (!currentMail?.itemId) {
    setResult('メール ID を取得できていません。対象メールを開いてから再実行してください。');
    return;
  }

  const body = await getBodyIfNeeded(settings.includeBody);
  const payload = buildPayload({ body });
  renderRequestPreview({ body });
  setResult('Flow へ送信中...');

  try {
    const headers = { 'Content-Type': 'application/json' };
    if (settings.authHeaderName && settings.authHeaderValue) {
      headers[settings.authHeaderName] = settings.authHeaderValue;
    }

    const response = await fetch(settings.flowUrl, {
      method: 'POST',
      headers,
      body: JSON.stringify(payload),
    });

    const text = await response.text();
    let data;
    try {
      data = JSON.parse(text);
    } catch {
      data = { raw: text };
    }

    if (!response.ok) {
      throw new Error(data?.error || data?.message || `HTTP ${response.status}`);
    }

    const lines = [];
    if (data.agent) lines.push(`エージェント: ${data.agent}`);
    if (data.classification) lines.push(`分類: ${data.classification}`);
    if (data.summary) lines.push('', '要約:', data.summary);
    if (data.replyDraft) lines.push('', '返信案:', data.replyDraft);
    if (!lines.length) lines.push(JSON.stringify(data, null, 2));
    setResult(lines.join('\n'));
  } catch (error) {
    setResult(`エラー: ${error.message}`);
  }
}

async function copyResult() {
  const text = $('result').textContent;
  await navigator.clipboard.writeText(text);
  setResult(`${text}\n\n---\n結果をクリップボードへコピーしました。`);
}

Office.onReady(async () => {
  applySettings();
  $('saveSettingsBtn').addEventListener('click', saveSettings);
  $('refreshBtn').addEventListener('click', readCurrentMail);
  $('runBtn').addEventListener('click', runFlow);
  $('copyBtn').addEventListener('click', copyResult);
  ['agentName','mode','flowUrl','authHeaderName','authHeaderValue','includeBody'].forEach((id) => {
    const el = $(id);
    el.addEventListener(id === 'mode' || id === 'includeBody' ? 'change' : 'input', () => renderRequestPreview());
  });
  await readCurrentMail();
});
