import type { RequestOptions } from '@@/plugin-request/request';
import type { RequestConfig } from '@umijs/max';
import { getIntl } from '@umijs/max';
import { message, notification } from 'antd';

// 错误处理方案： 错误类型
enum ErrorShowType {
  SILENT = 0,
  WARN_MESSAGE = 1,
  ERROR_MESSAGE = 2,
  NOTIFICATION = 3,
  REDIRECT = 9,
}

// Flask 后端常见响应：{ success, data?, error?: { code, message } | string, need_login? }
interface ResponseStructure {
  success?: boolean;
  data?: unknown;
  errorCode?: number | string;
  errorMessage?: string;
  showType?: ErrorShowType;
  error?:
    | string
    | {
        code?: string;
        message?: string;
        message_en?: string;
      };
  need_login?: boolean;
  message?: string;
}

function extractBizMessage(res: ResponseStructure): string {
  if (typeof res.error === 'string' && res.error) return res.error;
  if (res.error && typeof res.error === 'object') {
    return (
      res.error.message ||
      res.error.message_en ||
      res.error.code ||
      res.errorMessage ||
      '请求失败'
    );
  }
  return res.errorMessage || res.message || '请求失败';
}

/**
 * @name 错误处理
 * 适配 OutlookEmail Flask API（success/error 契约 + session 401）
 * @doc https://umijs.org/docs/max/request#配置
 */
export const errorConfig: RequestConfig = {
  errorConfig: {
    errorThrower: (res) => {
      const body = res as unknown as ResponseStructure;
      // 部分接口可能没有 success 字段；仅当显式 success===false 时抛业务错
      if (body && body.success === false) {
        const errorMessage = extractBizMessage(body);
        const errorCode =
          (typeof body.error === 'object' && body.error?.code) ||
          body.errorCode ||
          'BIZ_ERROR';
        const error: any = new Error(errorMessage);
        error.name = 'BizError';
        error.info = {
          errorCode,
          errorMessage,
          showType: body.need_login
            ? ErrorShowType.REDIRECT
            : (body.showType ?? ErrorShowType.ERROR_MESSAGE),
          data: body.data,
        };
        throw error;
      }
    },
    errorHandler: (error: any, opts: any) => {
      if (opts?.skipErrorHandler) throw error;

      // Flask 未登录
      const status = error?.response?.status;
      const data = error?.response?.data as ResponseStructure | undefined;
      if (
        status === 401 ||
        data?.need_login ||
        (typeof data?.error === 'object' && data.error?.code === 'AUTH_REQUIRED')
      ) {
        if (window.location.pathname !== '/ops/user/login') {
          const redirect = encodeURIComponent(
            window.location.pathname +
              window.location.search +
              window.location.hash,
          );
          window.location.href = `/ops/user/login?redirect=${redirect}`;
        }
        return;
      }

      if (error.name === 'BizError') {
        const errorInfo = error.info as ResponseStructure | undefined;
        if (errorInfo) {
          const { errorMessage, errorCode } = errorInfo as any;
          switch (errorInfo.showType) {
            case ErrorShowType.SILENT:
              break;
            case ErrorShowType.WARN_MESSAGE:
              message.warning(errorMessage);
              break;
            case ErrorShowType.ERROR_MESSAGE:
              message.error(errorMessage);
              break;
            case ErrorShowType.NOTIFICATION:
              notification.open({
                title: String(errorCode ?? ''),
                description: errorMessage,
              });
              break;
            case ErrorShowType.REDIRECT:
              window.location.href = '/ops/user/login';
              break;
            default:
              message.error(errorMessage);
          }
        }
        return;
      }

      if (error.response) {
        const msg = data ? extractBizMessage(data) : '';
        message.error(msg || `Response status:${error.response.status}`);
      } else if (typeof navigator !== 'undefined' && !navigator.onLine) {
        message.error(
          getIntl().formatMessage({
            id: 'app.request.offline',
            defaultMessage:
              'Network unavailable. Please check your connection and try again.',
          }),
        );
      } else if (error.request) {
        message.error('None response! Please retry.');
      } else {
        message.error('Request error, please retry.');
      }
    },
  },

  // 请求拦截器：始终携带 cookie（Flask session）
  requestInterceptors: [
    (config: RequestOptions) => {
      return {
        ...config,
        credentials: 'include',
        withCredentials: true,
      };
    },
  ],

  responseInterceptors: [],
};
