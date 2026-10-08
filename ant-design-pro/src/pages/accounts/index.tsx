import {
  DeleteOutlined,
  EditOutlined,
  ExportOutlined,
  KeyOutlined,
  LoadingOutlined,
  MailOutlined,
  PlusOutlined,
  ReloadOutlined,
  SyncOutlined,
  TagsOutlined,
} from '@ant-design/icons';
import {
  ModalForm,
  PageContainer,
  ProFormDigit,
  ProFormSelect,
  ProFormSwitch,
  ProFormText,
  ProFormTextArea,
  ProTable,
  type ActionType,
  type ProColumns,
} from '@ant-design/pro-components';
import { useQuery } from '@tanstack/react-query';
import { useIntl } from '@umijs/max';
import {
  App,
  Button,
  Dropdown,
  Form,
  Input,
  Modal,
  Popconfirm,
  Select,
  Space,
  Switch,
  Tag,
  Typography,
} from 'antd';
import React, { useMemo, useRef, useState } from 'react';
import {
  batchDeleteAccounts,
  batchNotificationToggle,
  batchUpdateAccountGroup,
  batchUpdateAccountStatus,
  createAccount,
  deleteAccount,
  downloadBlob,
  exportSelectedGroups,
  fetchAccount,
  fetchAccounts,
  fetchProviders,
  pickAccountErrorMessage,
  refreshAllAccounts,
  refreshSelectedAccounts,
  toggleAccountTelegram,
  updateAccount,
  verifyExportPassword,
  type AccountItem,
  type RefreshAllEvent,
  type RefreshSelectedEvent,
} from '@/services/outlook/accounts';
import {
  fetchGroups,
  isTempMailboxGroup,
  type GroupItem,
} from '@/services/outlook/groups';
import { extractEmailVerification } from '@/services/outlook/emails';
import {
  batchManageAccountTags,
  createTag,
  deleteTag as deleteTagApi,
  fetchTags,
  serializeTagIds,
  type TagItem,
} from '@/services/outlook/tags';
import {
  ACCOUNT_STATUS_OPTIONS,
  accountStatusLabel,
  refreshStatusLabel,
} from '@/utils/statusLabels';

const statusColor = (status?: string) => {
  const s = String(status || '').toLowerCase();
  if (s === 'active') return 'success';
  if (s === 'expired' || s === 'inactive' || s === 'disabled') return 'default';
  if (s === 'error' || s === 'failed') return 'error';
  return 'processing';
};

/** 把后端 tags 字段归一化成 TagItem[]（兼容字符串形态） */
const normalizeTags = (
  tags?: Array<{ id?: number; name?: string; color?: string } | string>,
): TagItem[] =>
  (tags || []).map((t) =>
    typeof t === 'string' ? { id: 0, name: t } : {
      id: Number(t.id || 0),
      name: String(t.name || ''),
      color: t.color,
    },
  );

const AccountsPage: React.FC = () => {
  const { message, modal } = App.useApp();
  const intl = useIntl();
  const actionRef = useRef<ActionType>(null);
  const [groupId, setGroupId] = useState<number | undefined>(undefined);
  const [importOpen, setImportOpen] = useState(false);
  const [editing, setEditing] = useState<AccountItem | null>(null);
  const [editAccountType, setEditAccountType] = useState<string>('outlook');
  const [editLoading, setEditLoading] = useState(false);
  const [editForm] = Form.useForm();
  const [importProvider, setImportProvider] = useState<string>('outlook');
  const [selectedRows, setSelectedRows] = useState<AccountItem[]>([]);
  const [exportOpen, setExportOpen] = useState(false);
  const [exportGroupIds, setExportGroupIds] = useState<number[]>([]);
  const [exportPassword, setExportPassword] = useState('');
  const [exporting, setExporting] = useState(false);
  const [refreshAllRunning, setRefreshAllRunning] = useState(false);
  const refreshAllRunningRef = useRef(false);
  const [tagManagerOpen, setTagManagerOpen] = useState(false);
  const [newTagName, setNewTagName] = useState('');
  const [newTagColor, setNewTagColor] = useState('#1677ff');
  const [tagCreating, setTagCreating] = useState(false);
  // 行内验证码提取状态（ZER-660）：key 为账号 id
  const [extractingIds, setExtractingIds] = useState<Set<number>>(new Set());
  const [extractResult, setExtractResult] = useState<{
    email: string;
    code?: string;
    link?: string;
    raw?: any;
  } | null>(null);

  const groupsQuery = useQuery({
    queryKey: ['groups'],
    queryFn: fetchGroups,
  });
  const providersQuery = useQuery({
    queryKey: ['providers'],
    queryFn: fetchProviders,
  });
  const tagsQuery = useQuery({
    queryKey: ['tags'],
    queryFn: fetchTags,
  });
  const tags: TagItem[] = useMemo(
    () => tagsQuery.data?.tags || [],
    [tagsQuery.data],
  );
  const tagOptions = useMemo(
    () =>
      tags.map((t) => ({
        label: t.name,
        value: t.id,
      })),
    [tags],
  );

  const normalGroups = useMemo(() => {
    return (groupsQuery.data?.groups || []).filter(
      (g: GroupItem) => !isTempMailboxGroup(g),
    );
  }, [groupsQuery.data]);

  const groupOptions = useMemo(
    () =>
      normalGroups.map((g) => ({
        label: `${g.name} (${g.account_count || 0})`,
        value: g.id,
      })),
    [normalGroups],
  );

  const providerOptions = useMemo(() => {
    const list = providersQuery.data?.providers || [];
    if (!list.length) {
      return [
        { label: 'Outlook', value: 'outlook' },
        { label: '智能识别（混合导入）', value: 'auto' },
        { label: '自定义 IMAP', value: 'custom' },
      ];
    }
    return list.map((p) => ({
      label: p.label || p.key,
      value: p.key,
    }));
  }, [providersQuery.data]);

  const openEdit = async (row: AccountItem) => {
    setEditing(row);
    setEditAccountType(row.account_type || 'outlook');
    setEditLoading(true);
    try {
      const res = await fetchAccount(row.id);
      if (res?.success === false || !res.account) {
        message.error(pickAccountErrorMessage(res, '加载账号失败'));
        setEditing(null);
        return;
      }
      const acc = res.account;
      setEditAccountType(acc.account_type || 'outlook');
      editForm.setFieldsValue({
        email: acc.email,
        password: '',
        client_id: acc.client_id || '',
        refresh_token: '',
        group_id: acc.group_id || undefined,
        remark: acc.remark || '',
        status: acc.status || 'active',
        account_type: acc.account_type || 'outlook',
      });
    } catch (error: any) {
      message.error(error?.message || '加载账号失败');
      setEditing(null);
    } finally {
      setEditLoading(false);
    }
  };

  const columns: ProColumns<AccountItem>[] = [
    {
      title: '邮箱',
      dataIndex: 'email',
      copyable: true,
      ellipsis: true,
      fieldProps: { placeholder: '搜索邮箱' },
    },
    {
      title: '分组',
      dataIndex: 'group_name',
      search: false,
      render: (_, row) => (
        <Space size={6}>
          <span
            style={{
              width: 8,
              height: 8,
              borderRadius: '50%',
              background: row.group_color || '#666',
              display: 'inline-block',
            }}
          />
          {row.group_name || '--'}
        </Space>
      ),
    },
    {
      title: '状态',
      dataIndex: 'status',
      width: 110,
      search: false,
      render: (_, row) => (
        <Tag
          color={statusColor(row.status)}
          style={{ cursor: 'pointer' }}
          onClick={() => void onToggleStatus(row)}
        >
          {accountStatusLabel(row.status)}
        </Tag>
      ),
    },
    {
      title: '类型',
      dataIndex: 'account_type',
      width: 100,
      search: false,
      render: (_, row) => row.account_type || row.provider || '--',
    },
    {
      title: '通知',
      dataIndex: 'telegram_push_enabled',
      width: 90,
      search: false,
      render: (_, row) => (
        <Switch
          size="small"
          checked={!!(row.telegram_push_enabled ?? row.notification_enabled)}
          onChange={(checked) => void onToggleTelegram(row, checked)}
        />
      ),
    },
    {
      title: '标签',
      dataIndex: 'tags',
      width: 180,
      search: false,
      render: (_, row) => {
        const list = normalizeTags(row.tags);
        if (!list.length) return '--';
        return (
          <Space size={4} wrap>
            {list.map((t, idx) => (
              <Tag key={t.id || `${t.name}-${idx}`} color={t.color}>
                {t.name}
              </Tag>
            ))}
          </Space>
        );
      },
    },
    {
      title: '标签筛选',
      dataIndex: 'tag_ids',
      valueType: 'select',
      hideInTable: true,
      fieldProps: {
        mode: 'multiple',
        placeholder: '按标签筛选',
        options: tagOptions,
        loading: tagsQuery.isLoading,
      },
    },
    {
      title: '备注',
      dataIndex: 'remark',
      ellipsis: true,
      search: false,
      render: (v) => v || '--',
    },
    {
      title: '最近刷新',
      dataIndex: 'last_refresh_at',
      width: 180,
      search: false,
      render: (_, row) => (
        <Space direction="vertical" size={0}>
          <Typography.Text>{row.last_refresh_at || '--'}</Typography.Text>
          {row.last_refresh_status ? (
            <Typography.Text type="secondary" style={{ fontSize: 12 }}>
              {refreshStatusLabel(row.last_refresh_status)}
            </Typography.Text>
          ) : null}
        </Space>
      ),
    },
    {
      title: '操作',
      valueType: 'option',
      width: 400,
      render: (_, row) => [
        <Button
          key="extract-code"
          type="link"
          icon={
            extractingIds.has(row.id) ? <LoadingOutlined /> : <KeyOutlined />
          }
          disabled={extractingIds.has(row.id)}
          onClick={() => void onExtractCode(row)}
        >
          验证码
        </Button>,
        <Select
          key="move-group"
          size="small"
          placeholder="移组"
          style={{ width: 110 }}
          value={null}
          options={groupOptions}
          onChange={(gid) => {
            if (gid != null) void onMoveGroup(row, Number(gid));
          }}
        />,
        <Button
          key="mailbox"
          type="link"
          icon={<MailOutlined />}
          href={`/ops/mailbox?account=${encodeURIComponent(row.email)}`}
        >
          邮件
        </Button>,
        <Dropdown
          key="tags"
          menu={{
            items: tags.length
              ? tags.map((t) => {
                  const has = normalizeTags(row.tags).some(
                    (x) => x.id === t.id,
                  );
                  return {
                    key: String(t.id),
                    label: `${has ? '✓ ' : '+ '}${t.name}`,
                    onClick: () => void onToggleRowTag(row, t),
                  };
                })
              : [{ key: 'empty', label: '暂无标签，请先在标签管理中新建', disabled: true }],
          }}
        >
          <Button type="link" icon={<TagsOutlined />}>
            打标
          </Button>
        </Dropdown>,
        <Button
          key="edit"
          type="link"
          icon={<EditOutlined />}
          onClick={() => void openEdit(row)}
        >
          编辑
        </Button>,
        <Popconfirm
          key="delete"
          title="确认删除该账号？"
          onConfirm={async () => {
            try {
              const res = await deleteAccount(row.id);
              if (res?.success === false) {
                message.error(
                  pickAccountErrorMessage(res, res.message || '删除失败'),
                );
                return;
              }
              message.success('删除成功');
              actionRef.current?.reload();
            } catch (error: any) {
              message.error(error?.message || '删除失败');
            }
          }}
        >
          <Button type="link" danger icon={<DeleteOutlined />}>
            删除
          </Button>
        </Popconfirm>,
      ],
    },
  ];

  const selectedIds = selectedRows.map((r) => r.id);

  const onToggleStatus = async (row: AccountItem) => {
    const next =
      String(row.status || '').toLowerCase() === 'active' ? 'inactive' : 'active';
    try {
      const res = await batchUpdateAccountStatus([row.id], next);
      if (res?.success === false) {
        message.error(pickAccountErrorMessage(res, '状态更新失败'));
        return;
      }
      message.success(res.message || `已设为${accountStatusLabel(next)}`);
      actionRef.current?.reload();
    } catch (error: any) {
      message.error(error?.message || '状态更新失败');
    }
  };

  const onToggleTelegram = async (row: AccountItem, enabled: boolean) => {
    try {
      const res = await toggleAccountTelegram(row.id, enabled);
      if (res?.success === false) {
        message.error(pickAccountErrorMessage(res, '通知切换失败'));
        return;
      }
      message.success(res.message || (enabled ? '已开启通知' : '已关闭通知'));
      actionRef.current?.reload();
    } catch (error: any) {
      message.error(error?.message || '通知切换失败');
    }
  };

  // 行内打标：已有该标签则移除，没有则添加
  const onToggleRowTag = async (row: AccountItem, tag: TagItem) => {
    const has = normalizeTags(row.tags).some((t) => t.id === tag.id);
    try {
      const res = await batchManageAccountTags(
        [row.id],
        tag.id,
        has ? 'remove' : 'add',
      );
      if (res?.success === false) {
        message.error(pickAccountErrorMessage(res, '标签操作失败'));
        return;
      }
      message.success(res.message || (has ? '已移除标签' : '已添加标签'));
      actionRef.current?.reload();
    } catch (error: any) {
      message.error(error?.message || '标签操作失败');
    }
  };

  const runBatchTag = async (tagId: number, action: 'add' | 'remove') => {
    const tag = tags.find((t) => t.id === tagId);
    if (!tag) return;
    await runBatch(
      `${action === 'add' ? '为选中账号添加' : '从选中账号移除'}标签「${tag.name}」？`,
      async () => {
        const res = await batchManageAccountTags(selectedIds, tagId, action);
        if (res?.success === false) {
          throw new Error(pickAccountErrorMessage(res, '标签操作失败'));
        }
        message.success(res.message || '标签操作完成');
      },
    );
  };

  const onCreateTag = async () => {
    const name = newTagName.trim();
    if (!name) {
      message.warning('请输入标签名称');
      return;
    }
    setTagCreating(true);
    try {
      const res = await createTag(name, newTagColor);
      if (res?.success === false) {
        message.error(pickAccountErrorMessage(res, '标签创建失败'));
        return;
      }
      message.success(res.message || '标签创建成功');
      setNewTagName('');
      await tagsQuery.refetch();
    } catch (error: any) {
      message.error(error?.message || '标签创建失败');
    } finally {
      setTagCreating(false);
    }
  };

  const onDeleteTag = async (tagId: number) => {
    try {
      const res = await deleteTagApi(tagId);
      if (res?.success === false) {
        message.error(pickAccountErrorMessage(res, '标签删除失败'));
        return;
      }
      message.success(res.message || '标签已删除');
      await tagsQuery.refetch();
      actionRef.current?.reload();
    } catch (error: any) {
      message.error(error?.message || '标签删除失败');
    }
  };

  // 行内验证码提取（ZER-660）：与旧站简洁模式验证码 chip 对齐
  const onExtractCode = async (row: AccountItem) => {
    if (extractingIds.has(row.id)) return;
    setExtractingIds((prev) => new Set(prev).add(row.id));
    try {
      const res = await extractEmailVerification(row.email);
      if (res?.success === false || (!res?.data && !res?.message)) {
        message.error(res?.message || '未提取到验证码');
        return;
      }
      const data = res?.data || {};
      const code = data.verification_code || data.code || data.formatted || '';
      const link = data.verification_link || data.link || '';
      if (!code && !link) {
        message.warning(res?.message || '最新邮件中未识别到验证码');
        return;
      }
      setExtractResult({ email: row.email, code, link, raw: res });
    } catch (error: any) {
      message.error(error?.message || '验证码提取失败');
    } finally {
      setExtractingIds((prev) => {
        const next = new Set(prev);
        next.delete(row.id);
        return next;
      });
    }
  };

  // 行内快速移动分组（ZER-660）：对齐旧站简洁模式行内「移动分组」
  const onMoveGroup = async (row: AccountItem, gid: number) => {
    try {
      const res = await batchUpdateAccountGroup([row.id], gid);
      if (res?.success === false) {
        message.error(pickAccountErrorMessage(res, '移动分组失败'));
        return;
      }
      message.success(res.message || '已移动分组');
      actionRef.current?.reload();
    } catch (error: any) {
      message.error(error?.message || '移动分组失败');
    }
  };

  const runBatch = async (
    title: string,
    runner: () => Promise<void>,
  ) => {
    if (!selectedIds.length) {
      message.warning('请先勾选账号');
      return;
    }
    modal.confirm({
      title,
      onOk: async () => {
        try {
          await runner();
          setSelectedRows([]);
          actionRef.current?.reload();
        } catch (error: any) {
          message.error(error?.message || '批量操作失败');
        }
      },
    });
  };

  const onRefreshAll = () => {
    if (refreshAllRunningRef.current) {
      message.warning('全量 Token 刷新正在执行，请勿重复触发');
      return;
    }

    modal.confirm({
      title: '确认全量刷新 Token？',
      content:
        '此操作将刷新全部可刷新的 Outlook 账号 Token，并可能因账号数量和限流持续较长时间。执行期间请勿重复发起刷新。',
      okText: '立即全量刷新',
      cancelText: '取消',
      okButtonProps: { danger: true },
      onOk: async () => {
        if (refreshAllRunningRef.current) {
          message.warning('全量 Token 刷新正在执行，请勿重复触发');
          return;
        }

        refreshAllRunningRef.current = true;
        setRefreshAllRunning(true);
        const loadingKey = 'refresh-all-token';
        message.loading({
          content: '正在启动全量 Token 刷新…',
          key: loadingKey,
          duration: 0,
        });

        try {
          const res = await refreshAllAccounts({
            onEvent: (event: RefreshAllEvent) => {
              if (event.type === 'start') {
                message.loading({
                  content: `正在全量刷新 Token… 0 / ${event.total}（成功 0，失败 0，跳过 0）`,
                  key: loadingKey,
                  duration: 0,
                });
              } else if (event.type === 'progress') {
                message.loading({
                  content: `正在全量刷新 Token… ${event.current} / ${event.total}（成功 ${event.success_count || 0}，失败 ${event.failed_count || 0}，跳过 0）`,
                  key: loadingKey,
                  duration: 0,
                });
              } else if (event.type === 'delay') {
                message.loading({
                  content: `全量刷新限流等待约 ${Math.ceil(event.seconds)} 秒…`,
                  key: loadingKey,
                  duration: 0,
                });
              }
            },
          });

          actionRef.current?.reload();
          message.destroy(loadingKey);
          modal.confirm({
            title: '全量 Token 刷新已完成',
            content: `成功 ${res.success_count} 个，失败 ${res.failed_count} 个，跳过 ${res.skipped_count} 个。可前往刷新日志查看明细。`,
            okText: '查看刷新日志',
            cancelText: '留在当前页',
            onOk: () => {
              window.location.assign('/refresh-log');
            },
          });
        } catch (error: any) {
          message.error({
            content: error?.message || '全量 Token 刷新失败',
            key: loadingKey,
          });
        } finally {
          refreshAllRunningRef.current = false;
          setRefreshAllRunning(false);
        }
      },
    });
  };

  const onExport = async () => {
    if (!exportGroupIds.length) {
      message.warning('请选择要导出的分组');
      return;
    }
    if (!exportPassword) {
      message.warning('请输入登录密码以二次验证');
      return;
    }
    setExporting(true);
    try {
      const verify = await verifyExportPassword(exportPassword);
      if (verify?.success === false || !verify.verify_token) {
        message.error(pickAccountErrorMessage(verify, '密码验证失败'));
        return;
      }
      const { blob, filename } = await exportSelectedGroups(
        exportGroupIds,
        verify.verify_token,
      );
      downloadBlob(blob, filename);
      message.success('导出成功');
      setExportOpen(false);
      setExportPassword('');
    } catch (error: any) {
      message.error(error?.message || '导出失败');
    } finally {
      setExporting(false);
    }
  };

  const isImapEdit = String(editAccountType || '').toLowerCase() === 'imap';

  return (
    <PageContainer
      title={intl.formatMessage({
        id: 'outlook.accounts.title',
        defaultMessage: '账号管理',
      })}
      subTitle={intl.formatMessage({
        id: 'outlook.accounts.subtitle',
        defaultMessage: '管理已接入的邮箱账号',
      })}
      extra={
        <Space wrap>
          <Button
            icon={<ReloadOutlined />}
            onClick={() => actionRef.current?.reload()}
          >
            刷新
          </Button>
          <Button
            icon={<ExportOutlined />}
            onClick={() => {
              setExportGroupIds(groupOptions.map((g) => g.value));
              setExportOpen(true);
            }}
          >
            导出
          </Button>
          <Button
            danger
            icon={<SyncOutlined />}
            loading={refreshAllRunning}
            disabled={refreshAllRunning}
            onClick={onRefreshAll}
          >
            全量刷新 Token
          </Button>
          <Button
            type="primary"
            icon={<PlusOutlined />}
            onClick={() => {
              setImportProvider('outlook');
              setImportOpen(true);
            }}
          >
            添加账号
          </Button>
        </Space>
      }
    >
      <ProTable<AccountItem>
        rowKey="id"
        actionRef={actionRef}
        columns={columns}
        cardBordered
        options={{ density: true, reload: true, setting: true }}
        form={{ syncToUrl: false }}
        rowSelection={{
          selectedRowKeys: selectedIds,
          onChange: (_keys, rows) => setSelectedRows(rows),
        }}
        tableAlertOptionRender={() => (
          <Space wrap>
            <Button
              size="small"
              onClick={() =>
                void runBatch('批量设为正常？', async () => {
                  const res = await batchUpdateAccountStatus(
                    selectedIds,
                    'active',
                  );
                  if (res?.success === false) {
                    throw new Error(
                      pickAccountErrorMessage(res, '批量状态失败'),
                    );
                  }
                  message.success(res.message || '已批量激活');
                })
              }
            >
              批量启用
            </Button>
            <Button
              size="small"
              onClick={() =>
                void runBatch('批量设为停用？', async () => {
                  const res = await batchUpdateAccountStatus(
                    selectedIds,
                    'inactive',
                  );
                  if (res?.success === false) {
                    throw new Error(
                      pickAccountErrorMessage(res, '批量状态失败'),
                    );
                  }
                  message.success(res.message || '已批量停用');
                })
              }
            >
              批量停用
            </Button>
            <Button
              size="small"
              onClick={() =>
                void runBatch('批量开启通知？', async () => {
                  const res = await batchNotificationToggle(selectedIds, true);
                  if (res?.success === false) {
                    throw new Error(
                      pickAccountErrorMessage(res, '批量通知失败'),
                    );
                  }
                  message.success(res.message || '已开启通知');
                })
              }
            >
              开通知
            </Button>
            <Button
              size="small"
              onClick={() =>
                void runBatch('批量关闭通知？', async () => {
                  const res = await batchNotificationToggle(selectedIds, false);
                  if (res?.success === false) {
                    throw new Error(
                      pickAccountErrorMessage(res, '批量通知失败'),
                    );
                  }
                  message.success(res.message || '已关闭通知');
                })
              }
            >
              关通知
            </Button>
            <Select
              size="small"
              placeholder="添加标签"
              style={{ width: 140 }}
              options={tagOptions}
              value={null}
              onChange={(tagId: number | undefined) => {
                if (tagId != null) void runBatchTag(tagId, 'add');
              }}
            />
            <Select
              size="small"
              placeholder="移除标签"
              style={{ width: 140 }}
              options={tagOptions}
              value={null}
              onChange={(tagId: number | undefined) => {
                if (tagId != null) void runBatchTag(tagId, 'remove');
              }}
            />
            <Button
              size="small"
              icon={<SyncOutlined />}
              onClick={() =>
                void runBatch('刷新选中账号 Token？', async () => {
                  const loadingKey = 'refresh-selected-token';
                  message.loading({
                    content: '正在刷新 Token…',
                    key: loadingKey,
                    duration: 0,
                  });
                  try {
                    const res = await refreshSelectedAccounts(selectedIds, {
                      onEvent: (event: RefreshSelectedEvent) => {
                        if (event.type === 'start') {
                          message.loading({
                            content: `正在刷新 Token… 0 / ${event.total}`,
                            key: loadingKey,
                            duration: 0,
                          });
                        } else if (event.type === 'progress') {
                          const done =
                            event.result === 'processing'
                              ? Math.max(0, (event.current || 1) - 1)
                              : event.current;
                          message.loading({
                            content: `正在刷新 Token… ${done} / ${event.total}`,
                            key: loadingKey,
                            duration: 0,
                          });
                        } else if (event.type === 'delay') {
                          message.loading({
                            content: `限流等待约 ${Math.ceil(event.seconds)} 秒…`,
                            key: loadingKey,
                            duration: 0,
                          });
                        }
                      },
                    });
                    if (res.failed_count > 0) {
                      const detail =
                        res.failed_list
                          ?.slice(0, 5)
                          .map(
                            (f) =>
                              `${f.email || f.id || '?'}: ${f.error || '失败'}`,
                          )
                          .join('；') || '';
                      message.warning({
                        content: detail
                          ? `${res.message}（${detail}${
                              res.failed_list.length > 5 ? '…' : ''
                            }）`
                          : res.message,
                        key: loadingKey,
                      });
                    } else {
                      message.success({
                        content: res.message || 'Token 刷新完成',
                        key: loadingKey,
                      });
                    }
                  } catch (error) {
                    message.destroy(loadingKey);
                    throw error;
                  }
                })
              }
            >
              刷新 Token
            </Button>
            <Select
              size="small"
              placeholder="移到分组"
              style={{ width: 160 }}
              options={groupOptions}
              onChange={(gid) =>
                void runBatch(`批量移动到分组 #${gid}？`, async () => {
                  const res = await batchUpdateAccountGroup(selectedIds, gid);
                  if (res?.success === false) {
                    throw new Error(
                      pickAccountErrorMessage(res, '批量移组失败'),
                    );
                  }
                  message.success(res.message || '已移组');
                })
              }
            />
            <Button
              size="small"
              danger
              onClick={() =>
                void runBatch(`确认删除选中的 ${selectedIds.length} 个账号？`, async () => {
                  const res = await batchDeleteAccounts(selectedIds);
                  if (res?.success === false) {
                    throw new Error(
                      pickAccountErrorMessage(res, '批量删除失败'),
                    );
                  }
                  message.success(res.message || '已删除');
                })
              }
            >
              批量删除
            </Button>
          </Space>
        )}
        toolbar={{
          search: false,
          actions: [
            <Select
              key="group"
              allowClear
              placeholder="按分组筛选"
              style={{ width: 220 }}
              options={groupOptions}
              loading={groupsQuery.isLoading}
              value={groupId}
              onChange={(value) => {
                setGroupId(value);
                actionRef.current?.reload();
              }}
            />,
            <Button
              key="tag-manager"
              icon={<TagsOutlined />}
              onClick={() => setTagManagerOpen(true)}
            >
              标签管理
            </Button>,
          ],
        }}
        request={async (params) => {
          try {
            const res = await fetchAccounts({
              page: params.current || 1,
              page_size: params.pageSize || 20,
              search: (params.email as string) || undefined,
              group_id: groupId,
              tag_ids: serializeTagIds(
                Array.isArray(params.tag_ids) ? params.tag_ids : undefined,
              ),
              sort_by: 'refresh_time',
              sort_order: 'asc',
            });
            return {
              data: res.accounts || [],
              success: !!res.success,
              total: res.pagination?.total_count || 0,
            };
          } catch (error: any) {
            message.error(error?.message || '加载账号失败');
            return { data: [], success: false, total: 0 };
          }
        }}
        pagination={{
          defaultPageSize: 20,
          showSizeChanger: true,
        }}
      />

      <ModalForm
        title="添加 / 导入账号"
        open={importOpen}
        modalProps={{
          destroyOnHidden: true,
          onCancel: () => setImportOpen(false),
        }}
        onOpenChange={setImportOpen}
        initialValues={{
          provider: 'outlook',
          group_id: groupId,
          add_to_pool: false,
          duplicate_strategy: 'skip',
          imap_port: 993,
        }}
        onValuesChange={(changed) => {
          if (Object.hasOwn(changed, 'provider')) {
            setImportProvider(String(changed.provider || 'outlook'));
          }
        }}
        onFinish={async (values) => {
          try {
            const provider = String(values.provider || 'outlook').toLowerCase();
            const payload: Record<string, any> = {
              account_string: String(values.account_string || '').trim(),
              add_to_pool: !!values.add_to_pool,
            };
            if (!payload.account_string) {
              message.error('请输入账号信息');
              return false;
            }
            if (provider === 'auto') {
              payload.provider = 'auto';
              payload.group_id = null;
              payload.duplicate_strategy = values.duplicate_strategy || 'skip';
              if (values.imap_host) {
                payload.imap_host = values.imap_host;
                payload.imap_port = values.imap_port || 993;
              }
            } else {
              payload.provider = provider;
              payload.group_id = values.group_id;
              if (provider === 'custom') {
                if (values.imap_host) {
                  payload.imap_host = values.imap_host;
                  payload.imap_port = values.imap_port || 993;
                }
              } else if (provider !== 'outlook' && values.imap_host) {
                payload.imap_host = values.imap_host;
                payload.imap_port = values.imap_port || 993;
              }
            }

            const res = await createAccount(payload);
            if (res?.success === false) {
              message.error(pickAccountErrorMessage(res, '导入失败'));
              return false;
            }
            let msg = res.message || '导入成功';
            const summary = res.summary;
            if (summary?.mode === 'auto' && summary.by_provider) {
              const parts = Object.entries(summary.by_provider as Record<string, any>)
                .map(([prov, stats]) => {
                  const s = stats || {};
                  return `${prov}: 成功 ${s.imported || 0}${
                    s.skipped ? `, 跳过 ${s.skipped}` : ''
                  }${s.failed ? `, 失败 ${s.failed}` : ''}`;
                })
                .join('；');
              if (parts) msg = `${msg}（${parts}）`;
            }
            message.success(msg);
            actionRef.current?.reload();
            return true;
          } catch (error: any) {
            message.error(
              pickAccountErrorMessage(
                error?.response?.data,
                error?.message || '导入失败',
              ),
            );
            return false;
          }
        }}
      >
        <ProFormSelect
          name="provider"
          label="提供商"
          options={providerOptions}
          rules={[{ required: true, message: '请选择提供商' }]}
          fieldProps={{ loading: providersQuery.isLoading }}
        />
        {importProvider !== 'auto' ? (
          <ProFormSelect
            name="group_id"
            label="目标分组"
            options={groupOptions}
            rules={[{ required: true, message: '请选择分组' }]}
            fieldProps={{ loading: groupsQuery.isLoading }}
          />
        ) : (
          <ProFormSelect
            name="duplicate_strategy"
            label="重复账号策略"
            options={[
              { label: '跳过', value: 'skip' },
              { label: '覆盖', value: 'overwrite' },
            ]}
          />
        )}
        <ProFormTextArea
          name="account_string"
          label="账号文本"
          placeholder={
            importProvider === 'outlook'
              ? '每行一个：邮箱----密码----client_id----refresh_token'
              : importProvider === 'auto'
                ? '支持混合格式，每行一个账号…\nOutlook: 邮箱----密码----client_id----refresh_token\nIMAP: 邮箱----授权码----provider\n临时邮箱: 仅邮箱地址'
                : '每行一个：邮箱----IMAP授权码/应用密码'
          }
          rules={[{ required: true, message: '请输入账号信息' }]}
          fieldProps={{ rows: 8 }}
        />
        {(importProvider === 'custom' ||
          importProvider === 'auto' ||
          (importProvider !== 'outlook' && importProvider !== 'auto')) && (
          <>
            <ProFormText
              name="imap_host"
              label={
                importProvider === 'auto'
                  ? '备用 IMAP 主机（可选）'
                  : 'IMAP 主机'
              }
              placeholder={
                importProvider === 'custom'
                  ? '必填，或在文本每行内嵌主机/端口'
                  : '可选'
              }
            />
            <ProFormDigit
              name="imap_port"
              label="IMAP 端口"
              min={1}
              max={65535}
              fieldProps={{ precision: 0 }}
            />
          </>
        )}
        <ProFormSwitch name="add_to_pool" label="同时加入邮箱池" />
      </ModalForm>

      <Modal
        title={`编辑账号 · ${editing?.email || ''}`}
        open={!!editing}
        confirmLoading={editLoading}
        destroyOnHidden
        width={640}
        onCancel={() => {
          setEditing(null);
          setEditAccountType('outlook');
          editForm.resetFields();
        }}
        onOk={async () => {
          if (!editing) return;
          try {
            const values = await editForm.validateFields();
            const body: Record<string, any> = {
              email: String(values.email || '').trim(),
              group_id: values.group_id,
              remark: values.remark || '',
              status: values.status || 'active',
            };
            if (values.password) body.password = values.password;
            if (values.client_id) body.client_id = values.client_id;
            if (values.refresh_token) body.refresh_token = values.refresh_token;

            const res = await updateAccount(editing.id, body);
            if (res?.success === false) {
              message.error(pickAccountErrorMessage(res, '更新失败'));
              return;
            }
            message.success(res.message || '账号更新成功');
            setEditing(null);
            editForm.resetFields();
            actionRef.current?.reload();
          } catch (error: any) {
            if (error?.errorFields) return;
            message.error(
              pickAccountErrorMessage(
                error?.response?.data,
                error?.message || '更新失败',
              ),
            );
          }
        }}
      >
        <Form form={editForm} layout="vertical" disabled={editLoading}>
          <Form.Item name="account_type" hidden>
            <Input />
          </Form.Item>
          <Form.Item
            name="email"
            label="邮箱"
            rules={[{ required: true, message: '邮箱不能为空' }]}
          >
            <Input />
          </Form.Item>
          <Form.Item
            name="password"
            label={isImapEdit ? '授权码 / 应用密码' : '密码'}
            extra="留空则不修改"
          >
            <Input.Password placeholder="留空则不修改" />
          </Form.Item>
          {!isImapEdit ? (
            <>
              <Form.Item name="client_id" label="Client ID">
                <Input />
              </Form.Item>
              <Form.Item
                name="refresh_token"
                label="Refresh Token"
                extra="留空则不修改；修改 Client ID 时必须同时填写"
              >
                <Input.Password placeholder="留空则不修改" />
              </Form.Item>
            </>
          ) : null}
          <Form.Item
            name="group_id"
            label="分组"
            rules={[{ required: true, message: '请选择分组' }]}
          >
            <Select options={groupOptions} loading={groupsQuery.isLoading} />
          </Form.Item>
          <Form.Item name="status" label="状态">
            <Select options={ACCOUNT_STATUS_OPTIONS} />
          </Form.Item>
          <Form.Item name="remark" label="备注">
            <Input.TextArea rows={2} maxLength={200} showCount />
          </Form.Item>
        </Form>
      </Modal>

      <Modal
        title="标签管理"
        open={tagManagerOpen}
        footer={null}
        onCancel={() => setTagManagerOpen(false)}
        width={520}
      >
        <Space direction="vertical" style={{ width: '100%' }} size={16}>
          <Space size={4} wrap>
            {tags.length ? (
              tags.map((t) => (
                <Popconfirm
                  key={t.id}
                  title={`确认删除标签「${t.name}」？将解除所有账号的关联。`}
                  onConfirm={() => void onDeleteTag(t.id)}
                >
                  <Tag color={t.color} style={{ cursor: 'pointer' }}>
                    {t.name} ×
                  </Tag>
                </Popconfirm>
              ))
            ) : (
              <Typography.Text type="secondary">暂无标签</Typography.Text>
            )}
          </Space>
          <Space.Compact style={{ width: '100%' }}>
            <Input
              placeholder="新标签名称"
              value={newTagName}
              maxLength={50}
              onChange={(e) => setNewTagName(e.target.value)}
              onPressEnter={() => void onCreateTag()}
              style={{ width: '60%' }}
            />
            <input
              type="color"
              aria-label="标签颜色"
              value={newTagColor}
              onChange={(e) => setNewTagColor(e.target.value)}
              style={{
                width: 48,
                height: 32,
                border: '1px solid #d9d9d9',
                borderRadius: 6,
                background: '#fff',
                padding: 2,
                cursor: 'pointer',
              }}
            />
            <Button
              type="primary"
              loading={tagCreating}
              onClick={() => void onCreateTag()}
            >
              新建标签
            </Button>
          </Space.Compact>
          <Typography.Text type="secondary" style={{ fontSize: 12 }}>
            点击标签可删除；在账号列表通过「打标」为账号添加或移除标签。
          </Typography.Text>
        </Space>
      </Modal>

      <Modal
        title={`验证码 · ${extractResult?.email || ''}`}
        open={!!extractResult}
        onCancel={() => setExtractResult(null)}
        footer={[
          <Button key="close" onClick={() => setExtractResult(null)}>
            关闭
          </Button>,
        ]}
      >
        <Space direction="vertical" style={{ width: '100%' }} size={12}>
          {extractResult?.code ? (
            <div>
              <Typography.Text type="secondary">验证码</Typography.Text>
              <Typography.Paragraph
                copyable={{ text: extractResult.code }}
                style={{ fontSize: 20, marginBottom: 0, fontWeight: 600 }}
              >
                {extractResult.code}
              </Typography.Paragraph>
            </div>
          ) : null}
          {extractResult?.link ? (
            <div>
              <Typography.Text type="secondary">验证链接</Typography.Text>
              <Typography.Paragraph
                copyable={{ text: extractResult.link }}
                ellipsis={{ rows: 2 }}
                style={{ marginBottom: 0 }}
              >
                {extractResult.link}
              </Typography.Paragraph>
            </div>
          ) : null}
          <Button
            type="link"
            icon={<MailOutlined />}
            href={`/ops/mailbox?account=${encodeURIComponent(
              extractResult?.email || '',
            )}`}
          >
            前往邮箱查看邮件
          </Button>
        </Space>
      </Modal>

      <Modal
        title="导出账号（二次验证）"
        open={exportOpen}
        confirmLoading={exporting}
        onCancel={() => {
          setExportOpen(false);
          setExportPassword('');
        }}
        onOk={() => void onExport()}
        okText="验证并导出"
      >
        <Space direction="vertical" style={{ width: '100%' }} size={12}>
          <div>
            <Typography.Text>选择分组</Typography.Text>
            <Select
              mode="multiple"
              style={{ width: '100%', marginTop: 8 }}
              options={groupOptions}
              value={exportGroupIds}
              onChange={setExportGroupIds}
              placeholder="选择要导出的分组"
            />
          </div>
          <div>
            <Typography.Text>登录密码</Typography.Text>
            <Input.Password
              style={{ marginTop: 8 }}
              value={exportPassword}
              onChange={(e) => setExportPassword(e.target.value)}
              placeholder="二次验证，防止误导出敏感凭据"
            />
          </div>
        </Space>
      </Modal>
    </PageContainer>
  );
};

export default AccountsPage;
