import { ensureCsrfToken, clearCsrfToken } from './auth';
import { outlookRequest } from './request';

export type AccountItem = {
  id: number;
  email: string;
  customer_secret?: string;
  customer_secret_enabled?: boolean;
  customer_secret_expires_at?: string | null;
  account_type?: string;
  provider?: string;
  client_id?: string;
  group_id?: number | null;
  group_name?: string;
  group_color?: string;
  remark?: string | null;
  status?: string;
  tags?: Array<{ id?: number; name?: string; color?: string } | string>;
  notification_enabled?: boolean;
  telegram_push_enabled?: boolean;
  last_refresh_at?: string | null;
  last_refresh_status?: string | null;
  last_refresh_error?: string | null;
  latest_email_subject?: string;
  latest_email_from?: string;
  latest_email_folder?: string;
  latest_email_received_at?: string;
  latest_verification_code?: string;
  created_at?: string;
  updated_at?: string;
};

export type AccountsResponse = {
  success: boolean;
  accounts: AccountItem[];
  pagination?: {
    page: number;
    page_size: number;
    total_count: number;
    total_pages: number;
  };
};

export type AccountListParams = {
  group_id?: number;
  page?: number;
  page_size?: number;
  search?: string;
  sort_by?: 'refresh_time' | 'email';
  sort_order?: 'asc' | 'desc';
  tag_id?: number | number[];
  /** 逗号分隔的标签 ID，与后端 /api/accounts 的 tag_ids 参数对应 */
  tag_ids?: string;
};

export async function fetchAccounts(params: AccountListParams = {}) {
  return outlookRequest<AccountsResponse>('/api/accounts', {
    method: 'GET',
    params,
  });
}

export async function deleteAccount(accountId: number) {
  return outlookRequest<{ success: boolean; message?: string; error?: any }>(
    `/api/accounts/${accountId}`,
    { method: 'DELETE' },
  );
}

export async function updateAccountRemark(accountId: number, remark: string) {
  return outlookRequest<{ success: boolean; message?: string; error?: any }>(
    `/api/accounts/${accountId}/remark`,
    {
      method: 'PATCH',
      data: { remark },
    },
  );
}

export async function createAccount(body: Record<string, any>) {
  return outlookRequest<{
    success: boolean;
    message?: string;
    message_en?: string;
    summary?: Record<string, any>;
    errors?: Array<Record<string, any>>;
    error?: any;
    [key: string]: any;
  }>('/api/accounts', {
    method: 'POST',
    data: body,
    skipErrorHandler: true,
  });
}

export async function fetchAccount(accountId: number) {
  return outlookRequest<{
    success: boolean;
    account?: AccountItem & {
      password?: string;
      client_id?: string;
      refresh_token?: string;
      has_password?: boolean;
      has_refresh_token?: boolean;
    };
    error?: any;
  }>(`/api/accounts/${accountId}`, {
    method: 'GET',
    skipErrorHandler: true,
  });
}

export async function resetCustomerSecret(accountId: number) {
  return outlookRequest<{
    success: boolean;
    customer_secret?: string;
    error?: any;
  }>(`/api/accounts/${accountId}/customer-secret/reset`, {
    method: 'POST',
    skipErrorHandler: true,
  });
}

export async function updateCustomerSecretAccess(
  accountId: number,
  enabled: boolean,
  expiresAt?: string | null,
) {
  return outlookRequest<{
    success: boolean;
    error?: any;
  }>(`/api/accounts/${accountId}/customer-secret`, {
    method: 'PATCH',
    data: {
      enabled,
      expires_at: expiresAt || null,
    },
    skipErrorHandler: true,
  });
}

export async function updateAccount(
  accountId: number,
  body: Record<string, any>,
) {
  return outlookRequest<{
    success: boolean;
    message?: string;
    message_en?: string;
    error?: any;
  }>(`/api/accounts/${accountId}`, {
    method: 'PUT',
    data: body,
    skipErrorHandler: true,
  });
}

export type ProviderItem = {
  key: string;
  label: string;
  account_type?: string;
  note?: string;
};

export async function fetchProviders() {
  return outlookRequest<{ success: boolean; providers?: ProviderItem[] }>(
    '/api/providers',
    { method: 'GET', skipErrorHandler: true },
  );
}

export function pickAccountErrorMessage(
  payload: any,
  fallback = '请求失败',
): string {
  if (!payload) return fallback;
  if (typeof payload.error === 'string' && payload.error) return payload.error;
  if (payload.error && typeof payload.error === 'object') {
    return (
      payload.error.message ||
      payload.error.message_en ||
      payload.error.code ||
      fallback
    );
  }
  if (typeof payload.message === 'string' && payload.message) {
    return payload.message;
  }
  if (Array.isArray(payload.errors) && payload.errors.length) {
    const first = payload.errors[0];
    const detail =
      typeof first === 'string'
        ? first
        : first?.error || first?.message || JSON.stringify(first);
    return `导入部分失败：${detail}`;
  }
  return fallback;
}

export async function batchUpdateAccountStatus(
  accountIds: number[],
  status: string,
) {
  return outlookRequest<{
    success: boolean;
    message?: string;
    updated_count?: number;
    failed_count?: number;
    error?: any;
  }>('/api/accounts/batch-update-status', {
    method: 'POST',
    data: { account_ids: accountIds, status },
    skipErrorHandler: true,
  });
}

export async function batchDeleteAccounts(accountIds: number[]) {
  return outlookRequest<{
    success: boolean;
    message?: string;
    deleted_count?: number;
    error?: any;
  }>('/api/accounts/batch-delete', {
    method: 'POST',
    data: { account_ids: accountIds },
    skipErrorHandler: true,
  });
}

export async function batchUpdateAccountGroup(
  accountIds: number[],
  groupId: number,
) {
  return outlookRequest<{
    success: boolean;
    message?: string;
    error?: any;
  }>('/api/accounts/batch-update-group', {
    method: 'POST',
    data: { account_ids: accountIds, group_id: groupId },
    skipErrorHandler: true,
  });
}

export async function batchNotificationToggle(
  accountIds: number[],
  enabled: boolean,
) {
  return outlookRequest<{
    success: boolean;
    message?: string;
    error?: any;
  }>('/api/accounts/batch-notification-toggle', {
    method: 'POST',
    data: { account_ids: accountIds, enabled },
    skipErrorHandler: true,
  });
}

export async function toggleAccountTelegram(
  accountId: number,
  enabled: boolean,
) {
  return outlookRequest<{
    success: boolean;
    message?: string;
    error?: any;
  }>(`/api/accounts/${accountId}/telegram-toggle`, {
    method: 'POST',
    data: { enabled },
    skipErrorHandler: true,
  });
}

/** SSE 事件：与后端 refresh_service.stream_refresh_selected_accounts 对齐 */
export type RefreshSelectedEvent =
  | {
      type: 'start';
      total: number;
      skipped_count?: number;
      delay_seconds?: number;
      run_id?: number;
      trace_id?: string;
      refresh_type?: string;
    }
  | {
      type: 'progress';
      current: number;
      total: number;
      email?: string;
      account_id?: number;
      result?: 'processing' | 'success' | 'failed' | string;
      error_message?: string | null;
      last_refresh_at?: string | null;
      success_count?: number;
      failed_count?: number;
    }
  | {
      type: 'delay';
      seconds: number;
    }
  | {
      type: 'complete';
      total: number;
      success_count: number;
      failed_count: number;
      failed_list?: Array<{ id?: number; email?: string; error?: string }>;
      invalid_token_failed_count?: number;
      invalid_token_failed_list?: Array<Record<string, any>>;
      run_id?: number;
    }
  | {
      type: 'error';
      error?: {
        code?: string;
        message?: string;
        message_en?: string;
        details?: any;
        trace_id?: string;
      };
    };

export type RefreshSelectedResult = {
  success: boolean;
  message: string;
  total: number;
  success_count: number;
  failed_count: number;
  skipped_count: number;
  failed_list: Array<{ id?: number; email?: string; error?: string }>;
  error?: RefreshSelectedEvent extends { type: 'error' }
    ? Extract<RefreshSelectedEvent, { type: 'error' }>['error']
    : any;
};

export type RefreshSelectedOptions = {
  signal?: AbortSignal;
  onEvent?: (event: RefreshSelectedEvent) => void;
};

/** 从 SSE 文本块中解析完整 `data:` 行（供单测与流式读取复用） */
export function parseSseDataLines(
  chunk: string,
  carry = '',
): { events: RefreshSelectedEvent[]; rest: string } {
  const buffer = `${carry}${chunk}`;
  const lines = buffer.split('\n');
  const rest = lines.pop() ?? '';
  const events: RefreshSelectedEvent[] = [];
  for (const line of lines) {
    const trimmed = line.trimEnd();
    if (!trimmed.startsWith('data:')) continue;
    const payload = trimmed.slice(5).trimStart();
    if (!payload || payload === '[DONE]') continue;
    try {
      events.push(JSON.parse(payload) as RefreshSelectedEvent);
    } catch {
      /* 忽略半截/脏行 */
    }
  }
  return { events, rest };
}

function summarizeRefreshComplete(
  event: Extract<RefreshSelectedEvent, { type: 'complete' }>,
  skippedCount: number,
): RefreshSelectedResult {
  const failedList = event.failed_list || [];
  const skippedHint =
    skippedCount > 0 ? `，跳过 ${skippedCount} 个不可刷新账号` : '';
  if ((event.failed_count || 0) === 0) {
    return {
      success: true,
      message: `Token 刷新完成：成功 ${event.success_count} 个${skippedHint}`,
      total: event.total,
      success_count: event.success_count,
      failed_count: event.failed_count,
      skipped_count: skippedCount,
      failed_list: failedList,
    };
  }
  return {
    success: true,
    message: `Token 刷新完成：成功 ${event.success_count} 个，失败 ${event.failed_count} 个${skippedHint}`,
    total: event.total,
    success_count: event.success_count,
    failed_count: event.failed_count,
    skipped_count: skippedCount,
    failed_list: failedList,
  };
}

async function postRefreshSelectedFetch(
  accountIds: number[],
  csrfToken: string | null,
  signal?: AbortSignal,
): Promise<Response> {
  const headers: Record<string, string> = {
    'Content-Type': 'application/json',
    Accept: 'text/event-stream',
  };
  if (csrfToken) {
    headers['X-CSRFToken'] = csrfToken;
  }
  return fetch('/api/accounts/refresh/selected', {
    method: 'POST',
    credentials: 'include',
    headers,
    body: JSON.stringify({ account_ids: accountIds }),
    signal,
  });
}

/**
 * 批量刷新选中账号 Token。
 * 后端返回 text/event-stream（start/progress/delay/complete/error），
 * 必须用 fetch 流式解析，不能走 umi/axios JSON 请求。
 */
export async function refreshSelectedAccounts(
  accountIds: number[],
  options: RefreshSelectedOptions = {},
): Promise<RefreshSelectedResult> {
  if (!accountIds.length) {
    throw new Error('请先勾选账号');
  }

  const OVERALL_TIMEOUT_MS = 120_000;
  const HEARTBEAT_TIMEOUT_MS = 30_000;
  const controller = new AbortController();
  const externalSignal = options.signal;
  const onExternalAbort = () => controller.abort();
  if (externalSignal) {
    if (externalSignal.aborted) {
      controller.abort();
    } else {
      externalSignal.addEventListener('abort', onExternalAbort, { once: true });
    }
  }

  let overallTimer: ReturnType<typeof setTimeout> | null = null;
  let heartbeatTimer: ReturnType<typeof setTimeout> | null = null;
  let timedOut = false;

  const clearTimers = () => {
    if (overallTimer) clearTimeout(overallTimer);
    if (heartbeatTimer) clearTimeout(heartbeatTimer);
    overallTimer = null;
    heartbeatTimer = null;
  };

  const armHeartbeat = () => {
    if (heartbeatTimer) clearTimeout(heartbeatTimer);
    heartbeatTimer = setTimeout(() => {
      timedOut = true;
      controller.abort();
    }, HEARTBEAT_TIMEOUT_MS);
  };

  overallTimer = setTimeout(() => {
    timedOut = true;
    controller.abort();
  }, OVERALL_TIMEOUT_MS);

  try {
    let csrf = await ensureCsrfToken();
    let response = await postRefreshSelectedFetch(
      accountIds,
      csrf,
      controller.signal,
    );

    // CSRF 失败时刷新 token 重试一次（与 outlookRequest 行为对齐）
    if (response.status === 400) {
      const peek = await response.clone().text();
      if (/csrf|CSRF/i.test(peek)) {
        clearCsrfToken();
        csrf = await ensureCsrfToken(true);
        response = await postRefreshSelectedFetch(
          accountIds,
          csrf,
          controller.signal,
        );
      } else {
        let payload: any = null;
        try {
          payload = JSON.parse(peek);
        } catch {
          /* ignore */
        }
        throw new Error(pickAccountErrorMessage(payload, 'Token 刷新失败'));
      }
    }

    if (!response.ok) {
      let payload: any = null;
      try {
        const text = await response.text();
        payload = JSON.parse(text);
      } catch {
        /* ignore */
      }
      throw new Error(pickAccountErrorMessage(payload, '刷新请求失败'));
    }

    if (!response.body) {
      throw new Error('刷新响应缺少流式正文');
    }

    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let carry = '';
    let skippedCount = 0;
    let sawTerminal = false;
    let completeResult: RefreshSelectedResult | null = null;

    armHeartbeat();

    while (!sawTerminal) {
      const { done, value } = await reader.read();
      if (done) break;

      armHeartbeat();
      const decoded = decoder.decode(value, { stream: true });
      const { events, rest } = parseSseDataLines(decoded, carry);
      carry = rest;

      for (const event of events) {
        options.onEvent?.(event);
        if (event.type === 'start') {
          skippedCount = event.skipped_count || 0;
        } else if (event.type === 'complete') {
          completeResult = summarizeRefreshComplete(event, skippedCount);
          sawTerminal = true;
          break;
        } else if (event.type === 'error') {
          const msg = pickAccountErrorMessage(
            { error: event.error },
            '刷新执行失败',
          );
          const err: any = new Error(msg);
          err.refreshError = event.error;
          throw err;
        }
      }
    }

    // 冲刷 decoder 尾部
    if (!sawTerminal && carry.trim()) {
      const { events } = parseSseDataLines(`${carry}\n`, '');
      for (const event of events) {
        options.onEvent?.(event);
        if (event.type === 'start') {
          skippedCount = event.skipped_count || 0;
        } else if (event.type === 'complete') {
          completeResult = summarizeRefreshComplete(event, skippedCount);
          sawTerminal = true;
        } else if (event.type === 'error') {
          const msg = pickAccountErrorMessage(
            { error: event.error },
            '刷新执行失败',
          );
          const err: any = new Error(msg);
          err.refreshError = event.error;
          throw err;
        }
      }
    }

    if (!completeResult) {
      throw new Error('刷新流已结束，但未收到完成事件');
    }
    return completeResult;
  } catch (error: any) {
    if (error?.name === 'AbortError') {
      throw new Error(
        timedOut
          ? '刷新请求超时，请检查网络或代理配置后重试'
          : '刷新请求已取消',
      );
    }
    throw error;
  } finally {
    clearTimers();
    if (externalSignal) {
      externalSignal.removeEventListener('abort', onExternalAbort);
    }
  }
}

export async function verifyExportPassword(password: string) {
  return outlookRequest<{
    success: boolean;
    verify_token?: string;
    message?: string;
    error?: any;
  }>('/api/export/verify', {
    method: 'POST',
    data: { password },
    skipErrorHandler: true,
  });
}

/** 兼容 axios / fetch Response：判断是否 2xx */
export function isHttpOk(response: any): boolean {
  if (!response || typeof response !== 'object') return false;
  if (typeof response.ok === 'boolean') return response.ok;
  const status = Number(response.status);
  if (Number.isFinite(status) && status > 0) {
    return status >= 200 && status < 300;
  }
  // 无 status/ok 时：仅当看起来像 axios/fetch 响应壳且带 data/headers 时保守放行
  if ('headers' in response || 'data' in response || 'body' in response) {
    return true;
  }
  return false;
}

/** 兼容 axios headers 对象与 fetch Headers.get() */
export function getResponseHeader(
  response: any,
  name: string,
): string | undefined {
  if (!response?.headers) return undefined;
  const headers = response.headers;
  if (typeof headers.get === 'function') {
    const viaGet =
      headers.get(name) ||
      headers.get(name.toLowerCase()) ||
      headers.get(name.toUpperCase());
    if (viaGet) return String(viaGet);
  }
  const lower = name.toLowerCase();
  const direct =
    headers[name] ??
    headers[lower] ??
    headers[name.toUpperCase()] ??
    headers['Content-Disposition'] ??
    headers['content-disposition'];
  if (Array.isArray(direct)) return direct[0] ? String(direct[0]) : undefined;
  if (direct != null) return String(direct);
  // axios 偶发 common 桶
  const common = headers.common;
  if (common && typeof common === 'object') {
    const fromCommon = common[name] ?? common[lower];
    if (fromCommon != null) return String(fromCommon);
  }
  return undefined;
}

/** 从 Content-Disposition 解析文件名 */
export function parseContentDispositionFilename(
  header?: string | null,
  fallback = `accounts_export_${Date.now()}.txt`,
): string {
  if (!header) return fallback;
  const star = header.match(/filename\*\s*=\s*UTF-8''([^;\n]+)/i);
  if (star?.[1]) {
    try {
      return decodeURIComponent(star[1].replace(/["']/g, '').trim());
    } catch {
      return star[1].replace(/["']/g, '').trim();
    }
  }
  const plain = header.match(/filename\s*=\s*("?)([^";\n]+)\1/i);
  if (plain?.[2]) {
    return plain[2].trim();
  }
  return fallback;
}

/** 导出选中分组（blob）；需先 verifyExportPassword 拿 token */
export async function exportSelectedGroups(
  groupIds: number[],
  verifyToken: string,
): Promise<{ blob: Blob; filename: string }> {
  const res = await outlookRequest<any>('/api/accounts/export-selected', {
    method: 'POST',
    data: { group_ids: groupIds },
    headers: { 'X-Export-Token': verifyToken },
    skipErrorHandler: true,
    getResponse: true,
    responseType: 'blob',
  } as any);

  // umi getResponse: true → { data, response }；也可能直接是 axios response / 裸 blob
  const hasWrappedResponse =
    res &&
    typeof res === 'object' &&
    'response' in res &&
    res.response != null;
  const response: any = hasWrappedResponse
    ? res.response
    : res &&
        typeof res === 'object' &&
        (typeof res.ok === 'boolean' ||
          typeof res.status === 'number' ||
          res.headers)
      ? res
      : null;
  const data = hasWrappedResponse
    ? res.data
    : res instanceof Blob
      ? res
      : (res?.data ?? res);

  if (response && !isHttpOk(response)) {
    let payload: any = null;
    try {
      if (data instanceof Blob) {
        const text = await data.text();
        payload = JSON.parse(text);
      } else if (typeof data === 'string') {
        payload = JSON.parse(data);
      } else if (data && typeof data === 'object') {
        payload = data;
      }
    } catch {
      /* ignore */
    }
    throw new Error(pickAccountErrorMessage(payload, '导出失败'));
  }

  if (!(data instanceof Blob) && (data == null || typeof data === 'object' && !(data instanceof ArrayBuffer))) {
    // 非 blob 且像错误 JSON
    if (data && typeof data === 'object' && data.success === false) {
      throw new Error(pickAccountErrorMessage(data, '导出失败'));
    }
  }

  const blob: Blob =
    data instanceof Blob
      ? data
      : new Blob([data], { type: 'text/plain;charset=utf-8' });

  const cd = getResponseHeader(response, 'Content-Disposition');
  const filename = parseContentDispositionFilename(cd);
  return { blob, filename };
}

export function downloadBlob(blob: Blob, filename: string) {
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  window.URL.revokeObjectURL(url);
  document.body.removeChild(a);
}


/** SSE 事件：与后端 refresh_service.stream_refresh_all_accounts 对齐。 */
export type RefreshAllEvent = RefreshSelectedEvent;

export type RefreshAllResult = {
  success: boolean;
  message: string;
  total: number;
  success_count: number;
  failed_count: number;
  skipped_count: number;
  failed_list: Array<{ id?: number; email?: string; error?: string }>;
};

export type RefreshAllOptions = {
  signal?: AbortSignal;
  onEvent?: (event: RefreshAllEvent) => void;
};

async function getRefreshAllFetch(
  csrfToken: string | null,
  signal?: AbortSignal,
): Promise<Response> {
  const headers: Record<string, string> = {
    Accept: 'text/event-stream',
  };
  if (csrfToken) {
    headers['X-CSRFToken'] = csrfToken;
  }
  return fetch('/api/accounts/refresh-all', {
    method: 'GET',
    credentials: 'include',
    headers,
    signal,
  });
}

function summarizeRefreshAllComplete(
  event: Extract<RefreshAllEvent, { type: 'complete' }>,
): RefreshAllResult {
  const skippedCount = Math.max(
    0,
    event.total - event.success_count - event.failed_count,
  );
  const summary = `全量 Token 刷新完成：成功 ${event.success_count} 个，失败 ${event.failed_count} 个，跳过 ${skippedCount} 个`;
  return {
    success: true,
    message: summary,
    total: event.total,
    success_count: event.success_count,
    failed_count: event.failed_count,
    skipped_count: skippedCount,
    failed_list: event.failed_list || [],
  };
}

/**
 * 全量刷新所有可刷新 Outlook 账号的 Token。
 * 后端返回 text/event-stream（start/progress/delay/complete/error），
 * 使用 fetch 流式解析以实时展示进度。全量任务的超时会在收到 start 事件后
 * 根据账号数量和限流间隔动态放宽，避免较大账号池被前端过早中断。
 */
export async function refreshAllAccounts(
  options: RefreshAllOptions = {},
): Promise<RefreshAllResult> {
  const BOOTSTRAP_TIMEOUT_MS = 60_000;
  const HEARTBEAT_TIMEOUT_MS = 90_000;
  const MAX_OVERALL_TIMEOUT_MS = 12 * 60 * 60_000;
  const controller = new AbortController();
  const externalSignal = options.signal;
  const onExternalAbort = () => controller.abort();
  if (externalSignal) {
    if (externalSignal.aborted) {
      controller.abort();
    } else {
      externalSignal.addEventListener('abort', onExternalAbort, { once: true });
    }
  }

  let overallTimer: ReturnType<typeof setTimeout> | null = null;
  let heartbeatTimer: ReturnType<typeof setTimeout> | null = null;
  let timedOut = false;

  const clearTimers = () => {
    if (overallTimer) clearTimeout(overallTimer);
    if (heartbeatTimer) clearTimeout(heartbeatTimer);
    overallTimer = null;
    heartbeatTimer = null;
  };

  const armOverall = (timeoutMs: number) => {
    if (overallTimer) clearTimeout(overallTimer);
    overallTimer = setTimeout(() => {
      timedOut = true;
      controller.abort();
    }, timeoutMs);
  };

  const armHeartbeat = () => {
    if (heartbeatTimer) clearTimeout(heartbeatTimer);
    heartbeatTimer = setTimeout(() => {
      timedOut = true;
      controller.abort();
    }, HEARTBEAT_TIMEOUT_MS);
  };

  armOverall(BOOTSTRAP_TIMEOUT_MS);

  try {
    let csrf = await ensureCsrfToken();
    let response = await getRefreshAllFetch(csrf, controller.signal);

    // 与标准请求保持一致：若服务端明确返回 CSRF 错误，则刷新令牌后重试一次。
    if (response.status === 400) {
      const peek = await response.clone().text();
      if (/csrf|CSRF/i.test(peek)) {
        clearCsrfToken();
        csrf = await ensureCsrfToken(true);
        response = await getRefreshAllFetch(csrf, controller.signal);
      } else {
        let payload: any = null;
        try {
          payload = JSON.parse(peek);
        } catch {
          /* ignore */
        }
        throw new Error(pickAccountErrorMessage(payload, '全量刷新请求失败'));
      }
    }

    if (!response.ok) {
      let payload: any = null;
      try {
        payload = JSON.parse(await response.text());
      } catch {
        /* ignore */
      }
      throw new Error(pickAccountErrorMessage(payload, '全量刷新请求失败'));
    }

    if (!response.body) {
      throw new Error('全量刷新响应缺少流式正文');
    }

    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let carry = '';
    let sawTerminal = false;
    let completeResult: RefreshAllResult | null = null;

    armHeartbeat();

    const handleEvent = (event: RefreshAllEvent) => {
      options.onEvent?.(event);
      if (event.type === 'start') {
        const delaySeconds = Math.max(0, event.delay_seconds || 0);
        const estimatedPerAccountMs = Math.max(15_000, (delaySeconds + 15) * 1_000);
        armOverall(
          Math.min(
            MAX_OVERALL_TIMEOUT_MS,
            Math.max(5 * 60_000, event.total * estimatedPerAccountMs),
          ),
        );
      } else if (event.type === 'complete') {
        completeResult = summarizeRefreshAllComplete(event);
        sawTerminal = true;
      } else if (event.type === 'error') {
        const msg = pickAccountErrorMessage(
          { error: event.error },
          '全量刷新执行失败',
        );
        const err: any = new Error(msg);
        err.refreshError = event.error;
        throw err;
      }
    };

    while (!sawTerminal) {
      const { done, value } = await reader.read();
      if (done) break;

      armHeartbeat();
      const decoded = decoder.decode(value, { stream: true });
      const { events, rest } = parseSseDataLines(decoded, carry);
      carry = rest;
      for (const event of events) {
        handleEvent(event);
        if (sawTerminal) break;
      }
    }

    if (!sawTerminal && carry.trim()) {
      const { events } = parseSseDataLines(`${carry}\n`, '');
      for (const event of events) {
        handleEvent(event);
        if (sawTerminal) break;
      }
    }

    if (!completeResult) {
      throw new Error('全量刷新流已结束，但未收到完成事件');
    }
    return completeResult;
  } catch (error: any) {
    if (error?.name === 'AbortError') {
      throw new Error(
        timedOut
          ? '全量刷新请求超时，请检查网络或刷新日志后重试'
          : '全量刷新请求已取消',
      );
    }
    throw error;
  } finally {
    clearTimers();
    if (externalSignal) {
      externalSignal.removeEventListener('abort', onExternalAbort);
    }
  }
}
