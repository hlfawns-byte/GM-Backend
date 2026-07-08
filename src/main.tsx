import React from "react";
import ReactDOM from "react-dom/client";
import * as XLSX from "xlsx";
import {
  Bell,
  ChevronDown,
  Database,
  Gift,
  KeyRound,
  LayoutDashboard,
  LogOut,
  Mail,
  MessageSquare,
  MonitorDot,
  Plus,
  Search,
  Send,
  Server,
  ShieldBan,
  ShieldCheck,
  SlidersHorizontal,
  Trash2,
  Trophy,
  UserCog,
  Users,
} from "lucide-react";
import "./styles.css";

type SectionKey =
  | "dashboard"
  | "playerInfo"
  | "uidAccount"
  | "accountUids"
  | "bindUid"
  | "gmState"
  | "silence"
  | "stateList"
  | "stateInfo"
  | "mailPersonal"
  | "mailGlobal"
  | "mailTemplate"
  | "mailRewardTemplate"
  | "giftCode"
  | "giftClaim"
  | "giftRecall"
  | "notice"
  | "sprint"
  | "systemMsg"
  | "serverTime"
  | "chatClear"
  | "rank"
  | "logs";

type Session = {
  adminAccount: string;
  operatorAccount: string;
  game: string;
  serverName: string;
  token: string;
  serverUrl: string;
  apiBase: string;
};

type ManagedAccount = {
  id: number;
  account: string;
  password?: string;
  displayName: string;
  avatarUrl?: string;
  role: string;
  games: string[];
  permissions: string[];
  isManager?: boolean;
  status: "启用" | "停用";
};

type AuthSession = {
  adminAccount: string;
  operatorAccount: string;
  displayName?: string;
  avatarUrl?: string;
  isAdmin: boolean;
  isRootAdmin: boolean;
  games: string[];
};

type GameConfig = {
  name: string;
  serverName: string;
  serverUrl: string;
  environment: string;
  logo: string;
  serverAccount?: string;
  serverPassword?: string;
  iconUrl?: string;
  backgroundUrl?: string;
  id?: number;
};

type ApiField = {
  key: string;
  label: string;
  placeholder?: string;
  defaultValue?: string;
  kind?: "text" | "number" | "textarea";
};

type ApiAction = {
  key: string;
  label: string;
  endpoint?: string;
  fields?: ApiField[];
  buildBody?: (values: Record<string, string>) => unknown;
};

type ModuleConfig = {
  title: string;
  description: string;
  icon: React.ComponentType<{ size?: number }>;
  status: "live" | "pending";
  actions: ApiAction[];
  listAction?: ApiAction;
};

type ApiResult = {
  id: number;
  label: string;
  endpoint: string;
  ok: boolean;
  status: number;
  payload: unknown;
};

type ApiPostResponse = {
  ok: boolean;
  status: number;
  payload: unknown;
};

type MailSectionKey = Extract<SectionKey, "mailPersonal" | "mailGlobal" | "mailTemplate" | "mailRewardTemplate">;
type GiftSectionKey = Extract<SectionKey, "giftCode" | "giftClaim" | "giftRecall">;

type ItemOption = {
  id: number;
  name: string;
  icon?: string;
};

type ServerOption = {
  id: number;
  name: string;
};

type MailTemplateContent = {
  title: string;
  body: string;
};

type MailTemplate = {
  id: string;
  name: string;
  title: string;
  body: string;
  contents?: Record<string, MailTemplateContent>;
  createdAt: string;
  updatedAt: string;
};

type RewardTemplate = {
  id: string;
  title: string;
  items: MailRewardItem[];
  createdAt: string;
  updatedAt: string;
};

type UserLog = {
  id: number;
  time: string;
  operator: string;
  role: string;
  action: string;
  target: string;
  game?: string;
  serverName?: string;
  detail?: string;
  result: "成功" | "失败";
};

type NoticeConfig = {
  slot: number;
  title: string;
  body: string;
  imagePath: string;
  regBegin?: string;
  regEnd?: string;
  platforms?: string;
  versions?: string;
  updatedAt?: string;
};

type MailRewardItem = {
  itemId: string;
  count: string;
};

const MAX_REWARD_COUNT = 2_000_000_000;
const MAX_REWARD_GROUPS = 100;
const MAX_TEMPLATE_NAME_LENGTH = 30;

function validateRewardRows(rewards: MailRewardItem[], items: ItemOption[]) {
  const normalized = rewards.map((reward) => ({ itemId: reward.itemId.trim(), count: reward.count.trim() }));
  const filled = normalized.filter((reward) => reward.itemId || reward.count);
  if (!filled.length) return rewards.length > 1 ? { ok: false, message: "ItemID不能为空", itemList: [] as number[] } : { ok: true, itemList: [] as number[] };
  const itemList: number[] = [];
  for (const reward of filled) {
    if (!reward.itemId) {
      return { ok: false, message: "ItemID不能为空", itemList: [] as number[] };
    }
    if (!reward.count || reward.count === "0") {
      return { ok: false, message: "ItemID数量应该>=1", itemList: [] as number[] };
    }
    const itemId = Number(reward.itemId);
    const count = Number(reward.count);
    if (!Number.isInteger(itemId) || itemId <= 0 || !Number.isInteger(count) || count <= 0) {
      return { ok: false, message: itemId > 0 ? "ItemID数量应该>=1" : "ItemID不能为空", itemList: [] as number[] };
    }
    if (count > MAX_REWARD_COUNT) {
      return { ok: false, message: `奖励数量不能超过 ${MAX_REWARD_COUNT}`, itemList: [] as number[] };
    }
    if (items.length > 0 && !items.some((item) => item.id === itemId)) {
      return { ok: false, message: `道具 ${itemId} 未匹配到Item表，不能保存`, itemList: [] as number[] };
    }
    itemList.push(itemId, count);
  }
  if (itemList.length > MAX_REWARD_GROUPS * 2) {
    return { ok: false, message: `奖励道具过多，请控制在${MAX_REWARD_GROUPS}组以内后再保存`, itemList: [] as number[] };
  }
  return { ok: true, itemList };
}

const MAIL_DEFAULT_REG_BEGIN = "2020-01-01T00:00";
const MAIL_DEFAULT_EXPIRE = "2050-12-31T23:59";

function defaultRegEndTime() {
  return toDatetimeLocal(new Date());
}

const mailLanguages = ["中文(简体)", "中文(繁体)", "英文", "韩文", "日文", "德语", "法语", "西班牙语", "葡萄牙语", "俄语", "意大利语", "印尼语", "泰语", "越南语"];
const defaultMailLanguage = mailLanguages[0];

const portal = import.meta.env.VITE_GM_PORTAL === "prod" ? "prod" : "test";
const portalGameConfig: GameConfig =
  portal === "prod"
    ? { name: "包包4", serverName: "正式服", serverUrl: "/gm-api", environment: "Prod", logo: "包" }
    : { name: "包包4", serverName: "测试服", serverUrl: "/gm-api", environment: "Test", logo: "包" };
const gameConfigs: GameConfig[] = [portalGameConfig];

const permissionOptions = ["用户查询", "名单管理", "邮件公告", "礼包码", "活动配置", "服务器工具", "日志审计"];

const navGroups = [
  { label: "总览", icon: LayoutDashboard, items: [{ key: "dashboard", label: "首页" }] },
  {
    label: "游戏用户",
    icon: Users,
    items: [
      { key: "playerInfo", label: "玩家信息" },
      { key: "uidAccount", label: "UID查账号" },
      { key: "accountUids", label: "账号查角色" },
      { key: "bindUid", label: "绑定账号UID" },
      { key: "gmState", label: "封号/解封" },
      { key: "silence", label: "禁言/解禁" },
      { key: "stateList", label: "名单清单" },
      { key: "stateInfo", label: "状态码列表" },
      { key: "rank", label: "排行榜查询" },
    ],
  },
  {
    label: "邮件管理",
    icon: Mail,
    items: [
      { key: "mailPersonal", label: "个人邮件" },
      { key: "mailGlobal", label: "全局邮件" },
      { key: "mailTemplate", label: "模板" },
      { key: "mailRewardTemplate", label: "奖励模板" },
    ],
  },
  {
    label: "礼包码",
    icon: Gift,
    items: [
      { key: "giftCode", label: "礼包码" },
      { key: "giftClaim", label: "领取记录" },
      { key: "giftRecall", label: "测试召回邮件" },
    ],
  },
  {
    label: "公告",
    icon: Bell,
    items: [
      { key: "notice", label: "公告" },
    ],
  },
  {
    label: "活动",
    icon: Trophy,
    items: [
      { key: "sprint", label: "冲刺活动" },
      { key: "systemMsg", label: "系统消息" },
    ],
  },
  {
    label: "系统",
    icon: SlidersHorizontal,
    items: [
      { key: "serverTime", label: "开服管理" },
      { key: "logs", label: "作业日志" },
    ],
  },
] satisfies Array<{
  label: string;
  icon: React.ComponentType<{ size?: number }>;
  items: Array<{ key: SectionKey; label: string }>;
}>;

const field = {
  userId: { key: "UserId", label: "用户ID", placeholder: "20190332", kind: "number" } satisfies ApiField,
  uid: { key: "Uid", label: "UID", placeholder: "20190332", kind: "number" } satisfies ApiField,
  accId: { key: "AccId", label: "玩家账号", placeholder: "test01" } satisfies ApiField,
  svrId: { key: "SvrId", label: "服务器ID", placeholder: "1", kind: "number" } satisfies ApiField,
};

const modules: Record<SectionKey, ModuleConfig> = {
  dashboard: { title: "首页", description: "GM后台运行概览", icon: LayoutDashboard, status: "live", actions: [] },
  playerInfo: {
    title: "玩家信息",
    description: "通过用户ID查询玩家信息",
    icon: Users,
    status: "live",
    actions: [{ key: "gmPlayerInfo", label: "查询玩家信息", endpoint: "/gmPlayerInfo", fields: [field.userId], buildBody: (v) => ({ Typ: 1, UserId: toNumberArray(v.UserId) }) }],
  },
  uidAccount: {
    title: "UID查账号",
    description: "通过玩家UID查询账号",
    icon: Search,
    status: "live",
    actions: [{ key: "gmGetAcountByUid", label: "查询账号", endpoint: "/gmGetAcountByUid", fields: [field.uid], buildBody: (v) => ({ Uid: Number(v.Uid) }) }],
  },
  accountUids: {
    title: "账号查角色",
    description: "通过账号查询角色UID列表",
    icon: Users,
    status: "live",
    actions: [{ key: "gmGetUidsByAcc", label: "查询角色", endpoint: "/gmGetUidsByAcc", fields: [field.accId], buildBody: (v) => ({ AccId: v.AccId }) }],
  },
  bindUid: {
    title: "绑定账号UID",
    description: "设置账号、服务器和UID绑定关系",
    icon: UserCog,
    status: "live",
    actions: [
      {
        key: "gmSetAcountUid",
        label: "设置UID",
        endpoint: "/gmSetAcountUid",
        fields: [field.svrId, { key: "UidSrc", label: "源UID", kind: "number" }, { key: "UidDst", label: "目标UID", kind: "number" }, field.accId],
        buildBody: (v) => ({ SvrId: Number(v.SvrId), UidSrc: Number(v.UidSrc), UidDst: Number(v.UidDst), AccId: v.AccId }),
      },
    ],
  },
  gmState: {
    title: "封号/解封",
    description: "对用户ID执行封号和解封操作",
    icon: ShieldBan,
    status: "live",
    actions: [{ key: "gmStateAdd", label: "提交名单状态", endpoint: "/gmStateAdd", fields: [field.userId, { key: "State", label: "状态", placeholder: "0/1/2/3/4", kind: "number" }], buildBody: (v) => ({ UserId: Number(v.UserId), State: Number(v.State) }) }],
  },
  silence: { title: "禁言/解禁", description: "禁言和解禁玩家聊天权限", icon: ShieldBan, status: "live", actions: [] },
  stateList: { title: "名单清单", description: "查看名单列表", icon: ShieldCheck, status: "live", actions: [{ key: "gmStateLst", label: "刷新名单", endpoint: "/gmStateLst", fields: [], buildBody: () => ({}) }] },
  stateInfo: { title: "状态码列表", description: "查看名单状态码说明", icon: Database, status: "live", actions: [{ key: "gmStateInfo", label: "读取状态码", endpoint: "/gmStateInfo", fields: [], buildBody: () => ({}) }] },
  mailPersonal: {
    title: "个人邮件",
    description: "按用户ID发送和管理个人邮件",
    icon: Mail,
    status: "live",
    actions: [
      { key: "gmMailLst", label: "邮件列表", endpoint: "/gmMailLst", fields: [], buildBody: () => ({}) },
      {
        key: "gmMailAdd",
        label: "发送邮件",
        endpoint: "/gmMailAdd",
        fields: [
          { key: "TargetID", label: "目标ID", placeholder: "20190332,20190333" },
          { key: "Titel", label: "标题" },
          { key: "Body", label: "正文", kind: "textarea" },
          { key: "ItemLst", label: "道具ID", placeholder: "101,200" },
        ],
        buildBody: (v) => ({ Typ: 1, TargetID: toNumberArray(v.TargetID), RegtBegin: 0, Regt: 0, Et: 1888888888, St: 0, SenderName: "GM", Titel: v.Titel, Body: v.Body, BodyData: [], BodyData2: [], ItemLst: toNumberArray(v.ItemLst), Platform: [], Version: [] }),
      },
      { key: "gmMailDel", label: "删除邮件", endpoint: "/gmMailDel", fields: [{ key: "Id", label: "邮件ID" }], buildBody: (v) => ({ Id: v.Id }) },
    ],
  },
  mailGlobal: { title: "全局邮件", description: "发送和管理全服邮件", icon: Mail, status: "live", actions: [] },
  mailTemplate: { title: "模板", description: "维护邮件标题和内容模板", icon: Mail, status: "live", actions: [] },
  mailRewardTemplate: { title: "奖励模板", description: "维护邮件奖励模板", icon: Gift, status: "live", actions: [] },
  giftCode: {
    title: "礼包码",
    description: "礼包码新增、查询、删除和本地存取",
    icon: Gift,
    status: "live",
    actions: [
      { key: "gmGiftLst", label: "礼包列表", endpoint: "/gmGiftLst", fields: [], buildBody: () => ({}) },
      { key: "gmGiftAdd", label: "新增礼包", endpoint: "/gmGiftAdd", fields: [{ key: "Name", label: "礼包名" }, { key: "Num", label: "数量", kind: "number" }, { key: "ItemLst", label: "道具ID", placeholder: "101,200" }], buildBody: (v) => ({ Name: v.Name, Num: Number(v.Num), ItemLst: toNumberArray(v.ItemLst) }) },
      { key: "gmGiftDel", label: "删除礼包", endpoint: "/gmGiftDel", fields: [{ key: "Id", label: "礼包ID" }], buildBody: (v) => ({ Id: v.Id }) },
      { key: "gmGiftSaveLocal", label: "保存本地", endpoint: "/gmGiftSaveLocal", fields: [], buildBody: () => ({}) },
      { key: "gmGiftLoadLocal", label: "读取本地", endpoint: "/gmGiftLoadLocal", fields: [], buildBody: () => ({}) },
    ],
  },
  giftClaim: { title: "领取记录", description: "礼包码领取记录查询", icon: Gift, status: "live", actions: [] },
  giftRecall: { title: "测试召回邮件", description: "测试召回邮件发送", icon: Mail, status: "live", actions: [] },
  notice: { title: "公告", description: "配置游戏内公告展示内容", icon: Bell, status: "live", actions: [] },
  sprint: { title: "冲刺活动", description: "冲刺活动新增和列表", icon: Trophy, status: "live", actions: [{ key: "gmSprintLst", label: "活动列表", endpoint: "/gmSprintLst", fields: [], buildBody: () => ({}) }, { key: "gmSprintAddLst", label: "新增活动", endpoint: "/gmSprintAddLst", fields: [{ key: "Body", label: "活动JSON", kind: "textarea" }], buildBody: (v) => parseJson(v.Body) }] },
  systemMsg: { title: "系统消息", description: "向在线玩家发送系统消息", icon: MessageSquare, status: "live", actions: [{ key: "gmSendSystemMsg", label: "发送消息", endpoint: "/gmSendSystemMsg", fields: [{ key: "Msg", label: "消息内容", kind: "textarea" }], buildBody: (v) => ({ Msg: v.Msg }) }] },
  serverTime: { title: "开服管理", description: "查询和配置开服时间", icon: Server, status: "live", actions: [] },
  chatClear: { title: "清理聊天", description: "清理聊天消息", icon: Send, status: "live", actions: [{ key: "gmClearChatMsg", label: "清理聊天", endpoint: "/gmClearChatMsg", fields: [], buildBody: () => ({}) }] },
  rank: { title: "排行榜查询", description: "当前接口文档暂未开放", icon: Trophy, status: "pending", actions: [] },
  logs: { title: "作业日志", description: "当前接口文档暂未开放", icon: Database, status: "pending", actions: [] },
};

function toNumberArray(value?: string) {
  return String(value ?? "")
    .split(/[\s,，;；]+/)
    .map((item) => Number(item.trim()))
    .filter((item) => Number.isFinite(item));
}

function toFlexibleNumberArray(value?: string) {
  return String(value ?? "")
    .split(/[\s,，;；]+/)
    .map((item) => parseFlexibleNumber(item.trim()))
    .filter((item) => Number.isFinite(item));
}

function toVersionNumberArray(value?: string) {
  return String(value ?? "")
    .split(/[\s,，;；]+/)
    .map((item) => item.trim())
    .filter(Boolean)
    .flatMap((item) => {
      if (!/^\d+(?:\.\d+){2,3}$/.test(item)) return [];
      const parts = item.split(".").map((part) => Number(part));
      if (!parts.every((part) => Number.isInteger(part) && part >= 0 && part <= 99)) return [];
      const versionCode = parts.reduce((total, part, index) => total + part * Math.pow(100, parts.length - index - 1), 0);
      return [versionCode, versionCode * 100 + 1];
    });
}

function toPlatformNumberArray(value?: string) {
  return String(value ?? "")
    .split(/[\s,，;；]+/)
    .map((item) => Number(item.trim()))
    .filter((item) => item === 1 || item === 2);
}

function dateToDatetimeLocal(value: string, endOfDay = false) {
  if (!value) return "";
  if (value.includes("T")) return value;
  return `${value}T${endOfDay ? "23:59" : "00:00"}`;
}

function templatePrimaryContent(template: MailTemplate) {
  const contents = template.contents ?? {};
  return contents[defaultMailLanguage] ?? Object.values(contents).find((content) => content.title?.trim() || content.body?.trim()) ?? { title: template.title, body: template.body };
}

function stringFromCell(value: unknown) {
  return String(value ?? "").trim();
}

async function parseMailTemplateFile(file: File) {
  const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: false });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) throw new Error("模板文件中没有可读取的工作表");
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "" });
  const usefulRows = rows.filter((row) => row.some((cell) => stringFromCell(cell)));
  if (!usefulRows.length) throw new Error("模板文件没有内容");
  const headerIndex = usefulRows.findIndex((row) => row.some((cell) => /标题|title/i.test(stringFromCell(cell))) && row.some((cell) => /内容|body|content/i.test(stringFromCell(cell))));
  const headers = headerIndex >= 0 ? usefulRows[headerIndex].map(stringFromCell) : [];
  const dataRows = usefulRows.slice(headerIndex >= 0 ? headerIndex + 1 : 0);
  const findColumn = (patterns: RegExp[], fallback: number) => {
    const index = headers.findIndex((header) => patterns.some((pattern) => pattern.test(header)));
    return index >= 0 ? index : fallback;
  };
  const languageColumn = findColumn([/语言/, /^lang(?:uage)?$/i], -1);
  const titleColumn = findColumn([/标题/, /^title$/i], 0);
  const bodyColumn = findColumn([/内容/, /^body$/i, /^content$/i], titleColumn === 0 ? 1 : 0);
  const nameColumn = findColumn([/模板/, /^name$/i], -1);
  const contents = Object.fromEntries(mailLanguages.map((language) => [language, { title: "", body: "" }])) as Record<string, MailTemplateContent>;
  if (languageColumn >= 0) {
    for (const item of dataRows) {
      const language = mailLanguages.find((candidate) => candidate === stringFromCell(item[languageColumn]));
      if (!language) continue;
      contents[language] = { title: stringFromCell(item[titleColumn]), body: stringFromCell(item[bodyColumn]) };
    }
  } else {
    for (const language of mailLanguages) {
      const titleIndex = headers.findIndex((header) => header.includes(language) && /标题|title/i.test(header));
      const bodyIndex = headers.findIndex((header) => header.includes(language) && /内容|body|content/i.test(header));
      if (titleIndex >= 0 || bodyIndex >= 0) {
        const first = dataRows[0] ?? [];
        contents[language] = { title: titleIndex >= 0 ? stringFromCell(first[titleIndex]) : "", body: bodyIndex >= 0 ? stringFromCell(first[bodyIndex]) : "" };
      }
    }
  }
  const row = dataRows.find((item) => stringFromCell(item[titleColumn]) || stringFromCell(item[bodyColumn]));
  if (row) {
    contents[defaultMailLanguage] = { title: stringFromCell(row[titleColumn]), body: stringFromCell(row[bodyColumn]) };
  }
  if (!Object.values(contents).some((content) => content.title || content.body)) throw new Error("模板文件没有邮件标题或邮件内容");
  const primary = contents[defaultMailLanguage].title || contents[defaultMailLanguage].body ? contents[defaultMailLanguage] : Object.values(contents).find((content) => content.title || content.body) ?? { title: "", body: "" };
  return {
    name: nameColumn >= 0 && row ? stringFromCell(row[nameColumn]) : "",
    title: primary.title,
    body: primary.body,
    contents,
  };
}

function parseFlexibleNumber(value: string) {
  if (/^\d+(?:\.\d+){2,3}$/.test(value)) {
    const parts = value.split(".").map((part) => Number(part));
    if (parts.every((part) => Number.isInteger(part) && part >= 0 && part <= 99)) {
      return parts.reduce((total, part, index) => total + part * Math.pow(100, parts.length - index - 1), 0);
    }
  }
  return Number(value);
}

function parseJson(value: string) {
  try {
    return value ? JSON.parse(value) : {};
  } catch {
    return {};
  }
}

function getObject(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" && !Array.isArray(value) ? (value as Record<string, unknown>) : null;
}

function getArray(value: unknown): unknown[] {
  return Array.isArray(value) ? value : [];
}

function getApiData(payload: unknown) {
  return getObject(getObject(payload)?.data) ?? getObject(payload);
}

function humanizeApiError(message: string) {
  if (/id\[\d+\].*服务器不存在/.test(message) || /服务器不存在/.test(message)) return "区服未填写，或填写的区服不存在";
  if (/TargetID|UserId|用户.*不存在|玩家.*不存在/i.test(message)) return "用户ID未填写，或用户不存在";
  if (/Platform|平台/i.test(message)) return "系统条件未填写，或选择的系统不支持";
  if (/Version|版本/i.test(message)) return "版本条件未填写，或版本格式不正确";
  if (/参数错误/.test(message)) return "条件参数填写不完整，请检查区服、系统、版本和时间";
  if (/Et.*当前时间.*10分钟/.test(message)) return "过期时间至少要比现在晚 10 分钟";
  if (/St.*大于.*0/.test(message)) return "当前接口暂不支持定时生效，请先用立即生效";
  if (/is not exist this key/i.test(message)) return "正式服没有查到这个 UID 对应的账号，请确认 UID 是否属于当前区服/正式服";
  return message;
}

function apiBusinessError(result: ApiPostResponse) {
  if (!result.ok) {
    const payloadError = String(getObject(result.payload)?.error ?? getObject(result.payload)?.result ?? getObject(result.payload)?.message ?? "");
    if (payloadError) return humanizeApiError(payloadError);
    if (result.status === 405) return "服务器拒绝了请求，请检查奖励数量、道具ID或接口是否支持当前参数";
    return `HTTP ${result.status}`;
  }
  const data = getApiData(result.payload);
  const resultCode = data?.Result ?? getObject(result.payload)?.Result;
  if (resultCode !== undefined && Number(resultCode) !== 0) {
    return humanizeApiError(String(data?.Desc ?? data?.Msg ?? getObject(result.payload)?.result ?? `Result ${resultCode}`));
  }
  return "";
}

function apiPayloadSummary(payload: unknown) {
  const payloadObject = getObject(payload);
  const data = getObject(payloadObject?.data) ?? payloadObject;
  const rawMessage = String(payloadObject?.result ?? data?.Desc ?? data?.Msg ?? "");
  const code = payloadObject?.code ?? payloadObject?.status;
  const result = data?.Result;
  if (rawMessage) return humanizeApiError(rawMessage);
  if (result !== undefined && Number(result) !== 0) return "接口返回失败，请确认查询条件是否存在于当前区服";
  if (code && Number(code) >= 400) return `服务端返回 ${String(code)}`;
  return "";
}

function formatCell(value: unknown) {
  if (value === null || value === undefined || value === "") return "暂无数据";
  if (Array.isArray(value)) return value.length ? value.join(", ") : "暂无数据";
  if (typeof value === "object") return JSON.stringify(value);
  return String(value);
}

function filledRewards(rewards: MailRewardItem[]) {
  return rewards
    .filter((reward) => reward.itemId.trim() || (reward.count.trim() && reward.count.trim() !== "0"))
    .map((reward) => ({ itemId: String(Number(reward.itemId)), count: String(Number(reward.count)) }));
}

function apiBaseFromServerUrl(url: string) {
  const normalized = url.trim();
  if (!normalized || normalized.startsWith("http")) return "/gm-api";
  return normalized.replace(/\/$/, "");
}

const storageKey = `touka-gm-session-${portal}`;

type StoredSession = {
  auth: AuthSession | null;
  session: Session | null;
  active: SectionKey;
};

function loadStoredSession(): StoredSession {
  try {
    const parsed = JSON.parse(localStorage.getItem(storageKey) ?? "{}") as Partial<StoredSession>;
    const auth = parsed.auth ? { ...parsed.auth, isRootAdmin: parsed.auth.isRootAdmin ?? parsed.auth.isAdmin } : null;
    return {
      auth,
      session: parsed.session ?? null,
      active: parsed.active && parsed.active in modules ? parsed.active : "dashboard",
    };
  } catch {
    return { auth: null, session: null, active: "dashboard" };
  }
}

function App() {
  const initialSession = React.useMemo(loadStoredSession, []);
  const [auth, setAuth] = React.useState<AuthSession | null>(initialSession.auth);
  const [session, setSession] = React.useState<Session | null>(initialSession.session);
  const [active, setActive] = React.useState<SectionKey>(initialSession.active);
  const [results, setResults] = React.useState<ApiResult[]>([]);
  const [accounts, setAccounts] = React.useState<ManagedAccount[]>([]);
  const [games, setGames] = React.useState<GameConfig[]>(gameConfigs);
  const [loadingAction, setLoadingAction] = React.useState("");
  const [accountPanelOpen, setAccountPanelOpen] = React.useState(false);
  const [gamePanelOpen, setGamePanelOpen] = React.useState(false);
  const [logPanelOpen, setLogPanelOpen] = React.useState(false);
  const [profilePanelOpen, setProfilePanelOpen] = React.useState(false);
  const [appBuild, setAppBuild] = React.useState("");
  const [versionNoticeOpen, setVersionNoticeOpen] = React.useState(false);
  const [openNavGroups, setOpenNavGroups] = React.useState<Record<string, boolean>>(() => Object.fromEntries(navGroups.map((group) => [group.label, true])));

  const refreshAccounts = React.useCallback(async () => {
    const response = await fetch("/local-api/accounts");
    const payload = (await response.json()) as { accounts?: ManagedAccount[] };
    setAccounts(payload.accounts ?? []);
  }, []);

  const refreshGames = React.useCallback(async () => {
    const response = await fetch("/local-api/games");
    const payload = (await response.json()) as { games?: GameConfig[] };
    setGames(payload.games?.length ? payload.games : gameConfigs);
  }, []);

  React.useEffect(() => {
    void refreshAccounts().catch(() => setAccounts([]));
    void refreshGames().catch(() => setGames(gameConfigs));
  }, [refreshAccounts, refreshGames]);

  React.useEffect(() => {
    localStorage.setItem(storageKey, JSON.stringify({ auth, session, active }));
  }, [auth, session, active]);

  React.useEffect(() => {
    let disposed = false;
    const checkVersion = async () => {
      try {
        const response = await fetch(`/local-api/app-version?t=${Date.now()}`, { cache: "no-store" });
        const payload = (await response.json().catch(() => ({}))) as { build?: string };
        const nextBuild = String(payload.build ?? "");
        if (!nextBuild || disposed) return;
        setAppBuild((currentBuild) => {
          if (!currentBuild) return nextBuild;
          if (currentBuild !== nextBuild) {
            setVersionNoticeOpen(true);
          }
          return currentBuild;
        });
      } catch {
        // Version polling should never interrupt normal GM operations.
      }
    };
    void checkVersion();
    const timer = window.setInterval(() => void checkVersion(), 60_000);
    return () => {
      disposed = true;
      window.clearInterval(timer);
    };
  }, []);

  const logout = () => {
    localStorage.removeItem(storageKey);
    setSession(null);
    setAuth(null);
    setActive("dashboard");
  };

  const reloginForNewVersion = () => {
    localStorage.removeItem(storageKey);
    window.location.reload();
  };

  const writeUserLog = React.useCallback(async (log: { action: string; target: string; detail?: string; result?: "成功" | "失败" }) => {
    const operator = session?.operatorAccount ?? auth?.operatorAccount ?? "unknown";
    await fetch("/local-api/user-logs", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        operator,
        role: auth?.isAdmin ? "管理员" : "成员",
        action: log.action,
        target: log.target,
        game: session?.game,
        serverName: session?.serverName,
        detail: log.detail,
        result: log.result ?? "成功",
      }),
    }).catch(() => undefined);
  }, [auth, session]);

  const versionModal = versionNoticeOpen ? <VersionUpdateModal onRelogin={reloginForNewVersion} /> : null;

  if (!auth) {
    return <><LoginScreen games={games} onLogin={setAuth} />{versionModal}</>;
  }

  const reorderGames = async (orderedIds: number[]) => {
    const response = await fetch("/local-api/games/reorder", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ids: orderedIds }),
    });
    if (!response.ok) {
      const payload = (await response.json().catch(() => ({}))) as { error?: string };
      throw new Error(payload.error || "保存游戏排序失败");
    }
    await refreshGames();
    void writeUserLog({ action: "调整游戏排序", target: `${orderedIds.length} 个游戏区服` });
  };

  if (!session) {
    return <><GameSelectScreen auth={auth} games={games} onEnter={setSession} onLogout={logout} onReorder={reorderGames} />{versionModal}</>;
  }

  const moduleConfig = modules[active];

  const refreshSessionToken = async () => {
    const response = await fetch("/local-api/gm-token", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ serverUrl: session.serverUrl }),
    });
    const payload = (await response.json().catch(() => ({}))) as { token?: string; adminAccount?: string; error?: string };
    if (!response.ok || !payload.token) throw new Error(payload.error || "Token 已过期，请重新登录");
    const nextSession = { ...session, adminAccount: payload.adminAccount ?? session.adminAccount, token: payload.token };
    setSession(nextSession);
    return payload.token;
  };

  const postWithToken = async (endpoint: string, body: unknown): Promise<ApiPostResponse> => {
    const send = async (token: string) => {
      const response = await fetch("/local-api/gm-post", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ serverUrl: session.serverUrl, endpoint, token, body }),
      });
      const text = await response.text();
      return { ok: response.ok, status: response.status, payload: parseJson(text) };
    };
    const first = await send(session.token);
    if (first.status !== 401) {
      void writeUserLog({ action: "服务端接口", target: endpoint, detail: JSON.stringify(body).slice(0, 300), result: first.ok ? "成功" : "失败" });
      return first;
    }
    const freshToken = await refreshSessionToken();
    const retried = await send(freshToken);
    void writeUserLog({ action: "服务端接口", target: endpoint, detail: JSON.stringify(body).slice(0, 300), result: retried.ok ? "成功" : "失败" });
    return retried;
  };

  const apiPost = async (action: ApiAction, values: Record<string, string>) => {
    if (!action.endpoint) return;
    setLoadingAction(action.key);
    try {
      const body = action.buildBody ? action.buildBody(values) : values;
      const result = await postWithToken(action.endpoint, body);
      setResults((current) => [{ id: Date.now(), label: action.label, endpoint: action.endpoint ?? "", ok: result.ok, status: result.status, payload: result.payload }, ...current].slice(0, 8));
    } finally {
      setLoadingAction("");
    }
  };

  const createAccount = async (account: ManagedAccount) => {
    const response = await fetch("/local-api/accounts", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(account) });
    if (!response.ok) {
      const payload = (await response.json().catch(() => ({}))) as { error?: string };
      throw new Error(payload.error || "创建账号失败");
    }
    await refreshAccounts();
    void writeUserLog({ action: "创建账号", target: account.account, detail: `${account.role} / ${(account.games ?? []).join(",")}` });
  };

  const updateAccount = async (account: ManagedAccount) => {
    const response = await fetch(`/local-api/accounts/${account.id}`, { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify(account) });
    if (!response.ok) {
      const payload = (await response.json().catch(() => ({}))) as { error?: string };
      throw new Error(payload.error || "保存账号失败");
    }
    await refreshAccounts();
    void writeUserLog({ action: "修改账号", target: account.account, detail: `${account.role} / ${(account.games ?? []).join(",")}` });
  };

  const deleteAccount = async (accountId: number) => {
    await fetch(`/local-api/accounts/${accountId}`, { method: "DELETE" });
    await refreshAccounts();
    void writeUserLog({ action: "删除账号", target: String(accountId) });
  };

  const createGame = async (game: GameConfig) => {
    const response = await fetch("/local-api/games", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(game) });
    if (!response.ok) {
      const payload = (await response.json().catch(() => ({}))) as { error?: string };
      throw new Error(payload.error || "新增区服失败");
    }
    await refreshGames();
    void writeUserLog({ action: "新增游戏区服", target: `${game.name}/${game.serverName}` });
  };

  const updateGame = async (game: GameConfig) => {
    if (!game.id) throw new Error("缺少区服ID");
    const response = await fetch(`/local-api/games/${game.id}`, { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify(game) });
    if (!response.ok) {
      const payload = (await response.json().catch(() => ({}))) as { error?: string };
      throw new Error(payload.error || "保存区服失败");
    }
    await refreshGames();
    await refreshAccounts();
    void writeUserLog({ action: "修改游戏区服", target: `${game.name}/${game.serverName}` });
  };

  const deleteGame = async (gameId: number) => {
    await fetch(`/local-api/games/${gameId}`, { method: "DELETE" });
    await refreshGames();
    await refreshAccounts();
    void writeUserLog({ action: "删除游戏区服", target: String(gameId) });
  };

  const updateProfile = async (profile: { displayName: string; avatarUrl: string; oldPassword?: string; newPassword?: string }) => {
    if (!auth) return;
    const response = await fetch("/local-api/profile", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ account: auth.operatorAccount, isRootAdmin: auth.isRootAdmin, ...profile }),
    });
    const payload = (await response.json().catch(() => ({}))) as { profile?: Partial<AuthSession>; error?: string };
    if (!response.ok) throw new Error(payload.error || "保存个人信息失败");
    const nextAuth = { ...auth, displayName: payload.profile?.displayName ?? profile.displayName, avatarUrl: payload.profile?.avatarUrl ?? profile.avatarUrl };
    setAuth(nextAuth);
    await refreshAccounts();
    void writeUserLog({ action: "修改个人信息", target: auth.operatorAccount });
  };

  return (
    <>
    <main className="app-shell">
      <aside className="sidebar">
        <div className="brand">
          <div className="brand-mark">T</div>
          <span className="brand-title">TOUKA GM</span>
        </div>
        <nav className="nav">
          {navGroups.map((group) => (
            <section key={group.label}>
              <button className="nav-group-title" onClick={() => setOpenNavGroups((current) => ({ ...current, [group.label]: !current[group.label] }))} type="button">
                <group.icon size={17} />
                {group.label}
                <ChevronDown className={`chevron ${openNavGroups[group.label] ? "open" : ""}`} size={15} />
              </button>
              {openNavGroups[group.label] && <div className="nav-items">
                {group.items.map((item) => (
                  <button className={`nav-item ${active === item.key ? "active" : ""}`} key={item.key} onClick={() => setActive(item.key)} type="button">
                    {item.label}
                    {modules[item.key].status === "pending" && <small>未开放</small>}
                  </button>
                ))}
              </div>}
            </section>
          ))}
        </nav>
      </aside>
      <section className="workspace">
        <header className="topbar">
          <div className="topbar-left">
            <span className="breadcrumb">{session.game} / {session.serverName} / {moduleConfig.title}</span>
            <span className="env-pill">{session.serverName}</span>
          </div>
          <div className="topbar-right">
            {auth.isAdmin && (
              <>
                <button className="account-button" onClick={() => setGamePanelOpen(true)} type="button"><Plus size={15} />游戏区服</button>
                <button className="account-button" onClick={() => setAccountPanelOpen(true)} type="button"><UserCog size={15} />账号管理</button>
                <button className="account-button" onClick={() => setLogPanelOpen(true)} type="button"><Database size={15} />用户日志</button>
              </>
            )}
            <button className="profile-button" onClick={() => setProfilePanelOpen(true)} type="button" title="个人信息设置">
              <AvatarImage account={session.operatorAccount} avatarUrl={auth.avatarUrl} displayName={auth.displayName || session.operatorAccount} />
              <span>{auth.displayName || session.operatorAccount}</span>
            </button>
            <button className="icon-button" onClick={() => setSession(null)} type="button" title="返回游戏选择"><LogOut size={18} /></button>
          </div>
        </header>
        <div className="content">
          {active !== "playerInfo" && (
            <section className="server-identity-banner">
              <MonitorDot size={22} />
              <div><span>当前游戏</span><b>{session.game} / {session.serverName}</b></div>
              <span className="token-status"><ShieldCheck size={15} />Token 已自动写入 Headers</span>
            </section>
          )}

          {active === "dashboard" ? <Dashboard results={results} accounts={accounts} /> : active === "playerInfo" ? <PlayerInfoPage postWithToken={postWithToken} /> : active === "bindUid" ? <BindUidPage postWithToken={postWithToken} /> : active === "gmState" ? <BanControlPage postWithToken={postWithToken} /> : active === "silence" ? <SilencePage operator={session.operatorAccount} /> : active.startsWith("mail") ? <MailSuitePage active={active as MailSectionKey} canUploadItemTable={auth.isAdmin} postWithToken={postWithToken} session={session} setActive={setActive} /> : active.startsWith("gift") ? <GiftSuitePage active={active as GiftSectionKey} canUploadItemTable={auth.isAdmin} postWithToken={postWithToken} /> : active === "notice" ? <NoticePage postWithToken={postWithToken} /> : active === "serverTime" ? <OpenServerPage postWithToken={postWithToken} /> : moduleConfig.status === "pending" ? <UnavailablePanel module={moduleConfig} /> : (
            <>
              <section className="filter-panel">
                <div className="panel-heading">
                  <div className="heading-title"><moduleConfig.icon size={20} /><strong>{moduleConfig.title}</strong></div>
                  <p>{moduleConfig.description}</p>
                </div>
              </section>
              <section className="action-grid">
                {moduleConfig.actions.map((action) => <ActionCard action={action} isLoading={loadingAction === action.key} key={action.key} onSubmit={(values) => void apiPost(action, values)} />)}
              </section>
              <ResultFeed results={results} />
            </>
          )}
        </div>
      </section>
      {auth.isAdmin && accountPanelOpen && <AccountPanel accounts={accounts} games={games} canManageAdmins={auth.isRootAdmin} session={session} onAdd={createAccount} onDelete={deleteAccount} onUpdate={updateAccount} onClose={() => setAccountPanelOpen(false)} />}
      {auth.isAdmin && gamePanelOpen && <GamePanel games={games} onAdd={createGame} onDelete={deleteGame} onUpdate={updateGame} onClose={() => setGamePanelOpen(false)} />}
      {auth.isAdmin && logPanelOpen && <UserLogPanel onClose={() => setLogPanelOpen(false)} />}
      {profilePanelOpen && auth && <ProfilePanel auth={auth} onClose={() => setProfilePanelOpen(false)} onSave={updateProfile} />}
    </main>
    {versionModal}
    </>
  );
}

function VersionUpdateModal({ onRelogin }: { onRelogin: () => void }) {
  return (
    <div className="modal-backdrop version-update-backdrop" role="presentation">
      <section className="version-update-modal" role="dialog" aria-modal="true" aria-labelledby="version-update-title">
        <div className="version-update-icon"><ShieldCheck size={26} /></div>
        <div>
          <h2 id="version-update-title">后台已更新</h2>
          <p>检测到当前GM后台有新版本，请重新登录后继续操作。</p>
        </div>
        <button className="primary-button" onClick={onRelogin} type="button"><LogOut size={16} />重新登录</button>
      </section>
    </div>
  );
}

function AvatarImage({ account, avatarUrl, displayName }: { account: string; avatarUrl?: string; displayName?: string }) {
  const label = (displayName || account || "U").trim().slice(0, 1).toUpperCase();
  return avatarUrl ? <img alt="" className="avatar-image" src={avatarUrl} /> : <span className="avatar-fallback">{label}</span>;
}

function ProfilePanel({ auth, onClose, onSave }: { auth: AuthSession; onClose: () => void; onSave: (profile: { displayName: string; avatarUrl: string; oldPassword?: string; newPassword?: string }) => Promise<void> }) {
  const [displayName, setDisplayName] = React.useState(auth.displayName || auth.operatorAccount);
  const [avatarUrl, setAvatarUrl] = React.useState(auth.avatarUrl || "");
  const [oldPassword, setOldPassword] = React.useState("");
  const [newPassword, setNewPassword] = React.useState("");
  const [confirmPassword, setConfirmPassword] = React.useState("");
  const [status, setStatus] = React.useState("");
  const [saving, setSaving] = React.useState(false);

  const uploadAvatar = (file?: File) => {
    if (!file) return;
    if (!file.type.startsWith("image/")) {
      setStatus("请选择图片文件");
      return;
    }
    if (file.size > 1024 * 1024) {
      setStatus("头像图片不能超过1MB");
      return;
    }
    const reader = new FileReader();
    reader.onload = () => setAvatarUrl(String(reader.result ?? ""));
    reader.onerror = () => setStatus("头像读取失败，请重新选择");
    reader.readAsDataURL(file);
  };

  const submit = async () => {
    if (!displayName.trim()) {
      setStatus("请输入显示名称");
      return;
    }
    if (oldPassword || newPassword || confirmPassword) {
      if (auth.isRootAdmin) {
        setStatus("admin 使用游戏服务端账号登录，密码请在游戏服务端侧维护");
        return;
      }
      if (!oldPassword) {
        setStatus("请输入原密码");
        return;
      }
      if (newPassword.length < 6) {
        setStatus("新密码至少6位");
        return;
      }
      if (newPassword !== confirmPassword) {
        setStatus("两次输入的新密码不一致");
        return;
      }
    }
    setSaving(true);
    setStatus("");
    try {
      await onSave({ displayName: displayName.trim(), avatarUrl, oldPassword: oldPassword || undefined, newPassword: newPassword || undefined });
      setStatus("个人信息已保存");
      setOldPassword("");
      setNewPassword("");
      setConfirmPassword("");
      setTimeout(onClose, 500);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : "保存失败");
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="modal-backdrop" role="presentation">
      <section className="profile-panel" role="dialog" aria-modal="true">
        <header><div><strong>个人信息设置</strong><span>{auth.operatorAccount}</span></div><button onClick={onClose} type="button">x</button></header>
        <div className="profile-form">
          <label className="profile-avatar-upload">
            <AvatarImage account={auth.operatorAccount} avatarUrl={avatarUrl} displayName={displayName} />
            <span>点击上传头像</span>
            <input accept="image/*" onChange={(event) => uploadAvatar(event.target.files?.[0])} type="file" />
          </label>
          <label>显示名称<input value={displayName} onChange={(event) => setDisplayName(event.target.value)} /></label>
          <div className="profile-password-block">
            <strong>修改密码</strong>
            {auth.isRootAdmin ? <p>admin 使用游戏服务端账号登录，密码请在游戏服务端侧维护。</p> : (
              <>
                <label>原密码<input autoComplete="current-password" type="password" value={oldPassword} onChange={(event) => setOldPassword(event.target.value)} /></label>
                <label>新密码<input autoComplete="new-password" type="password" value={newPassword} onChange={(event) => setNewPassword(event.target.value)} /></label>
                <label>确认新密码<input autoComplete="new-password" type="password" value={confirmPassword} onChange={(event) => setConfirmPassword(event.target.value)} /></label>
              </>
            )}
          </div>
          {status && <div className={status.includes("已保存") ? "profile-status success" : "profile-status"}>{status}</div>}
          <div className="profile-actions"><button disabled={saving} onClick={() => void submit()} type="button">{saving ? "保存中..." : "保存"}</button><button disabled={saving} onClick={onClose} type="button">取消</button></div>
        </div>
      </section>
    </div>
  );
}

function LoginScreen({ games, onLogin }: { games: GameConfig[]; onLogin: (session: AuthSession) => void }) {
  const [account, setAccount] = React.useState("");
  const [password, setPassword] = React.useState("");
  const [loginError, setLoginError] = React.useState("");
  const [isLoggingIn, setIsLoggingIn] = React.useState(false);

  return (
    <main className="login-page">
      <section className="login-panel">
        <div className="login-brand"><div className="brand-mark">T</div><div><strong>TOUKA GM 后台</strong><span>登录后选择可管理的游戏区服</span></div></div>
        <form className="login-form" onSubmit={async (event) => {
          event.preventDefault();
          const normalizedAccount = account.trim();
          if (!normalizedAccount || !password) {
            setLoginError("请输入账号和密码");
            return;
          }
          setIsLoggingIn(true);
          setLoginError("");
          try {
            let operator: ManagedAccount | undefined;
            let operatorAccount = normalizedAccount;
            let adminAccount = normalizedAccount;
            const operatorResponse = await fetch("/local-api/operator-login", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ account: normalizedAccount, password }) });
            if (operatorResponse.ok) {
              const operatorPayload = (await operatorResponse.json()) as { account?: ManagedAccount };
              operator = operatorPayload.account;
              operatorAccount = operatorPayload.account?.account ?? normalizedAccount;
              const tokenResponse = await fetch("/local-api/gm-token", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ serverUrl: portalGameConfig.serverUrl }),
              });
              const tokenPayload = (await tokenResponse.json().catch(() => ({}))) as { adminAccount?: string; error?: string };
              if (!tokenResponse.ok) throw new Error(tokenPayload.error || "账号或密码错误");
              adminAccount = tokenPayload.adminAccount ?? normalizedAccount;
            } else {
              const response = await fetch("/local-api/gm-login", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ account: normalizedAccount, password, serverUrl: portalGameConfig.serverUrl }),
              });
              const payload = (await response.json().catch(() => ({}))) as { adminAccount?: string; error?: string };
              if (!response.ok) throw new Error(payload.error || "账号或密码错误");
              adminAccount = payload.adminAccount ?? normalizedAccount;
            }
            let profile: Partial<AuthSession> = {};
            if (!operator) {
              const profileResponse = await fetch(`/local-api/profile?account=${encodeURIComponent(operatorAccount)}`);
              const profilePayload = (await profileResponse.json().catch(() => ({}))) as { profile?: Partial<AuthSession> };
              profile = profilePayload.profile ?? {};
            }
            onLogin({
              adminAccount,
              operatorAccount,
              displayName: operator?.displayName ?? profile.displayName ?? operatorAccount,
              avatarUrl: operator?.avatarUrl ?? profile.avatarUrl ?? "",
              isAdmin: !operator || Boolean(operator?.isManager),
              isRootAdmin: !operator,
              games: !operator || operator.isManager ? games.map((item) => `${item.name}/${item.serverName}`) : operator.games,
            });
            void fetch("/local-api/user-logs", {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({ operator: operatorAccount, role: operator ? "成员" : "管理员", action: "登录", target: "GM后台", result: "成功" }),
            });
          } catch (error) {
            void fetch("/local-api/user-logs", {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({ operator: normalizedAccount, role: "未知", action: "登录", target: "GM后台", result: "失败", detail: error instanceof Error ? error.message : "登录失败" }),
            });
            setLoginError(error instanceof Error ? error.message : "登录失败");
          } finally {
            setIsLoggingIn(false);
          }
        }}>
          <label>账号<input autoComplete="username" value={account} onChange={(event) => setAccount(event.target.value)} placeholder="请输入账号" /></label>
          <label>密码<input autoComplete="current-password" type="password" value={password} onChange={(event) => setPassword(event.target.value)} placeholder="请输入密码" /></label>
          <button className="login-submit" disabled={isLoggingIn} type="submit"><KeyRound size={17} />{isLoggingIn ? "正在登录" : "登录"}</button>
          {loginError && <div className="login-error">{loginError}</div>}
        </form>
        <div className="login-note"><ShieldCheck size={16} /><span>账号密码不预设、不展示；登录后只显示有权限的游戏区服。</span></div>
      </section>
    </main>
  );
}

function GameSelectScreen({ auth, games, onEnter, onLogout, onReorder }: { auth: AuthSession; games: GameConfig[]; onEnter: (session: Session) => void; onLogout: () => void; onReorder: (orderedIds: number[]) => Promise<void> }) {
  const [error, setError] = React.useState("");
  const [entering, setEntering] = React.useState("");
  const [orderedGames, setOrderedGames] = React.useState<GameConfig[]>([]);
  const [draggingId, setDraggingId] = React.useState<number | null>(null);
  const [pressingId, setPressingId] = React.useState<number | null>(null);
  const allowedGames = React.useMemo(() => games.filter((game) => auth.isAdmin || auth.games.includes(`${game.name}/${game.serverName}`)), [auth.games, auth.isAdmin, games]);
  const allowedSignature = React.useMemo(() => allowedGames.map((game) => game.id ?? `${game.name}/${game.serverName}`).join("|"), [allowedGames]);
  const pressTimerRef = React.useRef<number | null>(null);
  const draggedIdRef = React.useRef<number | null>(null);
  const dragMovedRef = React.useRef(false);
  const suppressClickRef = React.useRef(false);
  const latestOrderRef = React.useRef<GameConfig[]>(allowedGames);

  React.useEffect(() => {
    setOrderedGames(allowedGames);
    latestOrderRef.current = allowedGames;
  }, [allowedSignature, allowedGames]);

  const clearPressTimer = () => {
    if (pressTimerRef.current !== null) {
      window.clearTimeout(pressTimerRef.current);
      pressTimerRef.current = null;
    }
  };

  const beginPress = (event: React.PointerEvent<HTMLButtonElement>, game: GameConfig) => {
    if (!game.id || entering) return;
    const button = event.currentTarget;
    const pointerId = event.pointerId;
    clearPressTimer();
    setPressingId(game.id);
    pressTimerRef.current = window.setTimeout(() => {
      draggedIdRef.current = game.id ?? null;
      dragMovedRef.current = false;
      suppressClickRef.current = true;
      setDraggingId(game.id ?? null);
      setPressingId(null);
      button.setPointerCapture(pointerId);
    }, 420);
  };

  const moveDraggedGame = (targetId: number) => {
    const sourceId = draggedIdRef.current;
    if (!sourceId || sourceId === targetId) return;
    setOrderedGames((current) => {
      const fromIndex = current.findIndex((game) => game.id === sourceId);
      const toIndex = current.findIndex((game) => game.id === targetId);
      if (fromIndex < 0 || toIndex < 0 || fromIndex === toIndex) return current;
      const next = [...current];
      const [moved] = next.splice(fromIndex, 1);
      next.splice(toIndex, 0, moved);
      latestOrderRef.current = next;
      dragMovedRef.current = true;
      return next;
    });
  };

  const continueDrag = (event: React.PointerEvent<HTMLButtonElement>) => {
    if (!draggedIdRef.current) return;
    event.preventDefault();
    const target = document.elementFromPoint(event.clientX, event.clientY)?.closest("[data-game-id]");
    const targetId = Number(target?.getAttribute("data-game-id"));
    if (Number.isFinite(targetId)) moveDraggedGame(targetId);
  };

  const finishDrag = async (event?: React.PointerEvent<HTMLButtonElement>) => {
    clearPressTimer();
    setPressingId(null);
    const draggedId = draggedIdRef.current;
    draggedIdRef.current = null;
    setDraggingId(null);
    try {
      if (event?.currentTarget.hasPointerCapture(event.pointerId)) {
        event.currentTarget.releasePointerCapture(event.pointerId);
      }
    } catch {
      // Some browsers throw if capture was already released.
    }
    if (!draggedId || !dragMovedRef.current) {
      window.setTimeout(() => {
        suppressClickRef.current = false;
      }, 0);
      return;
    }
    const orderedIds = latestOrderRef.current.map((game) => game.id).filter((id): id is number => typeof id === "number");
    try {
      await onReorder(orderedIds);
    } catch (error) {
      setOrderedGames(allowedGames);
      latestOrderRef.current = allowedGames;
      setError(error instanceof Error ? error.message : "保存游戏排序失败");
    } finally {
      dragMovedRef.current = false;
      window.setTimeout(() => {
        suppressClickRef.current = false;
      }, 0);
    }
  };

  const enterGame = async (game: GameConfig) => {
    if (suppressClickRef.current) return;
    setEntering(`${game.name}/${game.serverName}`);
    setError("");
    try {
      const response = await fetch("/local-api/gm-token", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ serverUrl: game.serverUrl, serverAccount: game.serverAccount, serverPassword: game.serverPassword }),
      });
      const payload = (await response.json().catch(() => ({}))) as { token?: string; adminAccount?: string; error?: string };
      if (!response.ok || !payload.token) throw new Error(payload.error || "获取 Token 失败");
      onEnter({
        adminAccount: payload.adminAccount ?? auth.adminAccount,
        operatorAccount: auth.operatorAccount,
        game: game.name,
        serverName: game.serverName,
        token: payload.token,
        serverUrl: game.serverUrl,
        apiBase: apiBaseFromServerUrl(game.serverUrl),
      });
      void fetch("/local-api/user-logs", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ operator: auth.operatorAccount, role: auth.isAdmin ? "管理员" : "成员", action: "进入区服", target: `${game.name}/${game.serverName}`, game: game.name, serverName: game.serverName, result: "成功" }),
      });
    } catch (error) {
      void fetch("/local-api/user-logs", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ operator: auth.operatorAccount, role: auth.isAdmin ? "管理员" : "成员", action: "进入区服", target: `${game.name}/${game.serverName}`, game: game.name, serverName: game.serverName, result: "失败", detail: error instanceof Error ? error.message : "进入游戏失败" }),
      });
      setError(error instanceof Error ? error.message : "进入游戏失败");
    } finally {
      setEntering("");
    }
  };

  return (
    <main className="portal-shell">
      <aside className="portal-sidebar">
        <div className="portal-logo"><div className="brand-mark">T</div><strong>TOUKA</strong></div>
      </aside>
      <section className="portal-workspace">
        <header className="portal-topbar">
          <div className="topbar-left"><span className="breadcrumb">游戏列表</span></div>
          <div className="topbar-right">
            <div className="identity-stack"><span>{auth.operatorAccount}</span><small>请选择游戏区服</small></div>
            <button className="icon-button" onClick={onLogout} type="button" title="退出"><LogOut size={18} /></button>
          </div>
        </header>
        <div className="game-select-content">
          <h1>游戏列表</h1>
          {allowedGames.length === 0 ? (
            <section className="empty-game-list"><ShieldBan size={30} /><strong>暂无可访问游戏</strong><span>请联系管理员为该账号分配游戏区服权限。</span></section>
          ) : (
            <section className={`game-card-grid${draggingId ? " drag-active" : ""}`}>
              {orderedGames.map((game) => {
                const key = `${game.name}/${game.serverName}`;
                const isDragging = draggingId === game.id;
                const isPressing = pressingId === game.id;
                return (
                  <button
                    className={`game-card${isDragging ? " dragging" : ""}${isPressing ? " press-ready" : ""}`}
                    data-game-id={game.id}
                    disabled={entering === key}
                    key={game.id ?? key}
                    onClick={() => void enterGame(game)}
                    onPointerCancel={(event) => void finishDrag(event)}
                    onPointerDown={(event) => beginPress(event, game)}
                    onPointerLeave={clearPressTimer}
                    onPointerMove={continueDrag}
                    onPointerUp={(event) => void finishDrag(event)}
                    type="button"
                  >
                    <div className="game-cover" style={game.backgroundUrl ? { backgroundImage: `url(${game.backgroundUrl})` } : undefined}>
                      {!game.backgroundUrl && <span>{game.name}</span>}
                    </div>
                    <div className="game-card-body">
                      <div className="game-logo">{game.iconUrl ? <img alt={`${game.name} icon`} src={game.iconUrl} /> : game.logo}</div>
                      <div className="game-meta">
                        <strong>{game.name}</strong>
                        <span>{game.serverName}</span>
                      </div>
                      <small>{entering === key ? "进入中" : "进入后台"}</small>
                    </div>
                  </button>
                );
              })}
            </section>
          )}
          {error && <div className="form-error">{error}</div>}
        </div>
      </section>
    </main>
  );
}

function Dashboard({ results, accounts }: { results: ApiResult[]; accounts: ManagedAccount[] }) {
  const stats = [
    { label: "已接入接口", value: "18", delta: "POST + Token", icon: Database },
    { label: "账号数量", value: String(accounts.length), delta: "本机持久化", icon: Users },
    { label: "Token 状态", value: "有效", delta: "自动写入Header", icon: ShieldCheck },
    { label: "未开放模块", value: "3", delta: "UI已预留", icon: Bell },
  ];
  return (
    <>
      <section className="stats-grid">{stats.map((item) => <article className="stat-card" key={item.label}><item.icon size={22} /><span>{item.label}</span><strong>{item.value}</strong><small>{item.delta}</small></article>)}</section>
      <ResultFeed results={results} />
    </>
  );
}

function PlayerInfoPage({ postWithToken }: { postWithToken: (endpoint: string, body: unknown) => Promise<ApiPostResponse> }) {
  const [filters, setFilters] = React.useState({ userId: "", socialId: "", deviceId: "", orderId: "", nickname: "", accountId: "" });
  const [activeTab, setActiveTab] = React.useState("用户信息");
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState("");
  const [payload, setPayload] = React.useState<Record<string, unknown> | null>(null);

  const search = async () => {
    const ids = toNumberArray(filters.userId);
    if (ids.length === 0) {
      setError("当前接口只支持通过用户ID查询");
      setPayload(null);
      return;
    }
    setLoading(true);
    setError("");
    try {
      const result = await postWithToken("/gmPlayerInfo", { Typ: 1, UserId: ids });
      const parsed = result.payload as Record<string, unknown>;
      setPayload(parsed);
      if (!result.ok) setError(`服务端返回 ${result.status}`);
    } catch (error) {
      setError(error instanceof Error ? error.message : "查询失败");
      setPayload(null);
    } finally {
      setLoading(false);
    }
  };

  const data = getObject(payload?.data) ?? getObject(payload);
  const firstUserId = getArray(data?.UserId)[0] ?? filters.userId;
  const desc = typeof data?.Desc === "string" ? data.Desc : "";
  const summaryRows = data
    ? [{ userPid: firstUserId, deviceId: data.DeviceId, platformUid: data.AccId, registerTime: data.RegTime, lastLoginTime: data.LastLoginTime }]
    : [];
  const infoCells: Array<[string, unknown]> = [
    ["用户PID", firstUserId],
    ["deviceId", data?.DeviceId],
    ["platformUid", data?.AccId],
    ["金币", data?.Gold],
    ["钻石", data?.Diamond],
    ["等级", data?.Level],
    ["注册时间", data?.RegTime],
    ["最后登录时间", data?.LastLoginTime],
    ["国家", data?.Country],
    ["状态", desc || data?.State],
    ["是否付费", data?.IsPay],
    ["总充值金额", data?.PayTotal],
    ["总消耗钻石数量", data?.CostDiamond],
    ["总消耗金币数量", data?.CostGold],
    ["关卡ID", data?.StageId],
    ["系统标识", data?.System],
    ["客户端通信版本号", data?.ClientVersion],
    ["JSON", payload ? JSON.stringify(payload) : ""],
  ];
  const tabs = ["用户信息", "用户订单", "预下单", "用户行为", "封号日志", "禁赛日志", "月卡"];

  return (
    <section className="player-page">
      <div className="player-filter-bar">
        <PlayerSearchField label="用户ID" value={filters.userId} onChange={(value) => setFilters({ ...filters, userId: value })} placeholder="请输入用户Pid" />
        <PlayerSearchField disabled label="Social ID" value={filters.socialId} onChange={(value) => setFilters({ ...filters, socialId: value })} placeholder="当前接口暂未支持" />
        <PlayerSearchField disabled label="设备ID" value={filters.deviceId} onChange={(value) => setFilters({ ...filters, deviceId: value })} placeholder="当前接口暂未支持" />
        <PlayerSearchField disabled label="订单ID" value={filters.orderId} onChange={(value) => setFilters({ ...filters, orderId: value })} placeholder="当前接口暂未支持" />
        <PlayerSearchField disabled label="昵称" value={filters.nickname} onChange={(value) => setFilters({ ...filters, nickname: value })} placeholder="当前接口暂未支持" />
        <PlayerSearchField disabled label="账号唯一ID" value={filters.accountId} onChange={(value) => setFilters({ ...filters, accountId: value })} placeholder="当前接口暂未支持" />
        <button className="search-button" disabled={loading} onClick={() => void search()} type="button"><Search size={15} />{loading ? "查询中" : "搜索"}</button>
        {error && <div className="player-inline-error">{error}</div>}
      </div>

      <PlayerSimpleTable columns={["用户PID", "deviceId", "platformUid", "注册时间", "最后登录时间"]} emptyText={desc || "暂无数据"} rows={summaryRows} />

      <section className="player-detail-panel">
        <div className="player-tabs">
          {tabs.map((tab) => <button className={activeTab === tab ? "active" : ""} key={tab} onClick={() => setActiveTab(tab)} type="button">{tab}</button>)}
        </div>
        {activeTab === "用户信息" ? <PlayerInfoGrid cells={infoCells} /> : <div className="player-empty-block">当前服务器接口暂未开放“{activeTab}”数据</div>}
      </section>

      <PlayerSectionTable title="游戏角色" columns={["角色ID", "角色名称", "星级", "正在使用"]} />
      <PlayerSectionTable title="角色装备" columns={["装备ID", "装备名称", "等级", "正在使用"]} />
      <PlayerSectionTable title="角色符文" columns={["符文ID", "符文名称", "正在使用"]} />
    </section>
  );
}

function PlayerSearchField({ disabled, label, onChange, placeholder, value }: { disabled?: boolean; label: string; onChange: (value: string) => void; placeholder: string; value: string }) {
  return <label className="player-search-field"><span>{label}:</span><input disabled={disabled} value={value} onChange={(event) => onChange(event.target.value)} placeholder={placeholder} /></label>;
}

function PlayerSimpleTable({ columns, emptyText, rows }: { columns: string[]; emptyText: string; rows: Array<Record<string, unknown>> }) {
  const keys = ["userPid", "deviceId", "platformUid", "registerTime", "lastLoginTime"];
  return (
    <section className="player-table">
      <table><thead><tr>{columns.map((column) => <th key={column}>{column}</th>)}</tr></thead><tbody>{rows.length ? rows.map((row, index) => <tr key={index}>{keys.map((key) => <td key={key}>{formatCell(row[key])}</td>)}</tr>) : <tr><td colSpan={columns.length}>{emptyText}</td></tr>}</tbody></table>
    </section>
  );
}

function PlayerInfoGrid({ cells }: { cells: Array<[string, unknown]> }) {
  return <div className="player-info-grid">{cells.map(([label, value]) => <React.Fragment key={label}><div className="cell-label">{label}</div><div>{formatCell(value)}</div></React.Fragment>)}</div>;
}

function PlayerSectionTable({ columns, title }: { columns: string[]; title: string }) {
  return <section className="player-section-table"><div>{title}</div><PlayerSimpleTable columns={columns} emptyText="暂无数据" rows={[]} /></section>;
}

function BindUidPage({ postWithToken }: { postWithToken: (endpoint: string, body: unknown) => Promise<ApiPostResponse> }) {
  const [uid, setUid] = React.useState("");
  const [result, setResult] = React.useState<{ uid: string; account: string } | null>(null);
  const [modalOpen, setModalOpen] = React.useState(false);
  const [svrId, setSvrId] = React.useState("");
  const [newUid, setNewUid] = React.useState("");
  const [message, setMessage] = React.useState("");
  const [loading, setLoading] = React.useState(false);

  const search = async () => {
    if (!uid.trim()) {
      setMessage("请输入 UID");
      return;
    }
    setLoading(true);
    setMessage("");
    try {
      const response = await postWithToken("/gmGetAcountByUid", { Uid: Number(uid) });
      const data = getObject(getObject(response.payload)?.data) ?? getObject(response.payload);
      const account = String(data?.AccId ?? data?.Account ?? "");
      if (!response.ok || !account) {
        setResult(null);
        setMessage(response.ok ? "未查询到账号" : `查询失败：HTTP ${response.status}`);
        return;
      }
      setResult({ uid: String(data?.Uid ?? uid), account });
    } catch (error) {
      setResult(null);
      setMessage(error instanceof Error ? error.message : "查询失败");
    } finally {
      setLoading(false);
    }
  };

  const submit = async () => {
    if (!result) return;
    if (!svrId.trim() || !newUid.trim()) {
      setMessage("请输入服务器ID和新UID");
      return;
    }
    setLoading(true);
    setMessage("");
    try {
      const response = await postWithToken("/gmSetAcountUid", { SvrId: Number(svrId), UidSrc: Number(result.uid), UidDst: Number(newUid), AccId: result.account });
      setMessage(response.ok ? "设置成功" : `设置失败：HTTP ${response.status}`);
      if (response.ok) setModalOpen(false);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "设置失败");
    } finally {
      setLoading(false);
    }
  };

  return (
    <section className="bind-uid-page">
      <h2>设置用户账号</h2>
      <div className="bind-search-row">
        <input value={uid} onChange={(event) => setUid(event.target.value)} placeholder="请输入UID" />
        {uid && <button className="bind-clear" onClick={() => { setUid(""); setResult(null); }} type="button">×</button>}
        <button disabled={loading} onClick={() => void search()} type="button">{loading ? "查询中" : "查询"}</button>
      </div>
      {result && (
        <section className="bind-result-card">
          <span className="bind-watermark">admin</span>
          <h3>查询结果</h3>
          <div className="bind-result-body">
            <div className="bind-avatar">A</div>
            <div><strong>Account: {result.account}</strong><strong>Uid: {result.uid}</strong></div>
          </div>
          <button onClick={() => { setSvrId(""); setNewUid(""); setModalOpen(true); }} type="button"><UserCog size={15} />设置</button>
        </section>
      )}
      {message && <div className="bind-message">{message}</div>}
      {modalOpen && result && (
        <div className="modal-backdrop" role="presentation">
          <section className="bind-modal" role="dialog" aria-modal="true">
            <header><strong>设置</strong><button onClick={() => setModalOpen(false)} type="button">×</button></header>
            <div className="bind-modal-form">
              <label>原Uid<input disabled value={result.uid} /></label>
              <label>原账号ID<input disabled value={result.account} /></label>
              <label><span>*</span>服务器ID<input value={svrId} onChange={(event) => setSvrId(event.target.value)} placeholder="请输入服务器ID" /></label>
              <label><span>*</span>新UID<input value={newUid} onChange={(event) => setNewUid(event.target.value)} placeholder="请输入新的UID" /></label>
              <footer><button onClick={() => setModalOpen(false)} type="button">取消</button><button disabled={loading} onClick={() => void submit()} type="button">确定</button></footer>
            </div>
          </section>
        </div>
      )}
    </section>
  );
}

function BanControlPage({ postWithToken }: { postWithToken: (endpoint: string, body: unknown) => Promise<ApiPostResponse> }) {
  const [banTarget, setBanTarget] = React.useState<"device" | "account">("device");
  const [banMode, setBanMode] = React.useState<"temporary" | "permanent">("temporary");
  const [banDays, setBanDays] = React.useState("1");
  const [banUserIds, setBanUserIds] = React.useState("");
  const [unbanTarget, setUnbanTarget] = React.useState<"device" | "account">("device");
  const [unbanUserIds, setUnbanUserIds] = React.useState("");
  const [loading, setLoading] = React.useState<"ban" | "unban" | "">("");
  const [message, setMessage] = React.useState("");

  const idsFromText = (value: string) => value
    .split(/[\n,，\s]+/)
    .map((item) => Number(item.trim()))
    .filter((item) => Number.isFinite(item) && item > 0);

  const submitState = async (type: "ban" | "unban") => {
    const ids = idsFromText(type === "ban" ? banUserIds : unbanUserIds);
    if (!ids.length) {
      setMessage("请输入至少一个用户ID");
      return;
    }
    if (type === "ban" && banTarget === "device") {
      setMessage("当前服务器接口暂未开放设备封禁，仅支持账号封禁");
      return;
    }
    if (type === "ban" && banMode === "temporary") {
      setMessage("当前服务器接口暂未开放按天数临时封禁，将不会提交");
      return;
    }
    if (type === "unban" && unbanTarget === "device") {
      setMessage("当前服务器接口暂未开放设备解封，仅支持账号解封");
      return;
    }

    setLoading(type);
    setMessage("");
    try {
      const state = type === "ban" ? 4 : 0;
      const responses = await Promise.all(ids.map((userId) => postWithToken("/gmStateAdd", { UserId: userId, State: state })));
      const failed = responses.filter((response) => !response.ok || response.status >= 400);
      setMessage(failed.length ? `${type === "ban" ? "封号" : "解封"}完成，${failed.length}/${ids.length} 个请求失败` : `${type === "ban" ? "封号" : "解封"}提交成功，共 ${ids.length} 个用户ID`);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "提交失败");
    } finally {
      setLoading("");
    }
  };

  return (
    <section className="ban-page">
      <div className="ban-card">
        <section className="ban-form-block">
          <div className="ban-row">
            <span className="ban-label">封禁类型：</span>
            <RadioPill checked={banTarget === "device"} label="设备" onChange={() => setBanTarget("device")} />
            <RadioPill checked={banTarget === "account"} label="账号" onChange={() => setBanTarget("account")} />
          </div>
          <div className="ban-row">
            <span className="ban-label">封禁方式：</span>
            <RadioPill checked={banMode === "temporary"} label="暂时封禁" onChange={() => setBanMode("temporary")} />
            <RadioPill checked={banMode === "permanent"} label="永久封禁" onChange={() => setBanMode("permanent")} />
          </div>
          <label className="ban-days-row">
            <span className="ban-label">封禁天数：</span>
            <input value={banDays} onChange={(event) => setBanDays(event.target.value)} />
            <span>天</span>
          </label>
          <label className="ban-textarea-label">
            <span>userId：</span>
            <textarea value={banUserIds} onChange={(event) => setBanUserIds(event.target.value)} placeholder="请输入用户ID，一行一个用户ID" />
          </label>
          <div className="ban-actions">
            <button disabled={loading === "ban"} onClick={() => void submitState("ban")} type="button">{loading === "ban" ? "提交中" : "封号"}</button>
            <button onClick={() => setBanUserIds("")} type="button">取消</button>
          </div>
        </section>

        <section className="ban-form-block">
          <div className="ban-row">
            <span className="ban-label">解封类型：</span>
            <RadioPill checked={unbanTarget === "device"} label="设备" onChange={() => setUnbanTarget("device")} />
            <RadioPill checked={unbanTarget === "account"} label="账号" onChange={() => setUnbanTarget("account")} />
          </div>
          <label className="ban-textarea-label ban-unban-area">
            <span>userId：</span>
            <textarea value={unbanUserIds} onChange={(event) => setUnbanUserIds(event.target.value)} placeholder="请输入用户ID，一行一个用户ID" />
          </label>
          <div className="ban-actions">
            <button disabled={loading === "unban"} onClick={() => void submitState("unban")} type="button">{loading === "unban" ? "提交中" : "解封"}</button>
            <button onClick={() => setUnbanUserIds("")} type="button">取消</button>
          </div>
        </section>
      </div>
      {message && <div className="ban-status">{message}</div>}
    </section>
  );
}

function RadioPill({ checked, label, onChange }: { checked: boolean; label: string; onChange: () => void }) {
  return <label className="ban-radio"><input checked={checked} onChange={onChange} type="radio" />{label}</label>;
}

function SilencePage({ operator }: { operator: string }) {
  const [view, setView] = React.useState<"list" | "edit">("list");
  const [muteTypes, setMuteTypes] = React.useState(["末世回响"]);
  const [unmuteTypes, setUnmuteTypes] = React.useState(["末世回响"]);
  const [until, setUntil] = React.useState(toDatetimeLocal(new Date(Date.now() + 24 * 60 * 60 * 1000)));
  const [muteUserIds, setMuteUserIds] = React.useState("");
  const [unmuteUserIds, setUnmuteUserIds] = React.useState("");
  const [query, setQuery] = React.useState("");
  const [records, setRecords] = React.useState<Array<{ id: number; userId: string; until: string; status: string; createdAt: string; operator: string }>>([]);
  const silenceTypes = ["末世回响", "公会远征", "试炼之路"];

  const idsFromText = (value: string) => value.split(/[\n,，\s]+/).map((item) => item.trim()).filter(Boolean);
  const toggleType = (value: string, list: string[], setter: (items: string[]) => void) => {
    setter(list.includes(value) ? list.filter((item) => item !== value) : [...list, value]);
  };
  const addRecords = (type: "mute" | "unmute") => {
    const ids = idsFromText(type === "mute" ? muteUserIds : unmuteUserIds);
    if (!ids.length) return;
    const now = new Date().toISOString().slice(0, 19).replace("T", " ");
    setRecords((current) => [
      ...ids.map((userId) => ({ id: Date.now() + Number(userId.slice(-4) || 0), userId, until: type === "mute" ? formatBeijingTime(until) : "已解禁", status: type === "mute" ? "禁言" : "解禁", createdAt: now, operator })),
      ...current,
    ]);
    if (type === "mute") setMuteUserIds("");
    else setUnmuteUserIds("");
    setView("list");
  };

  const filtered = records.filter((record) => !query.trim() || record.userId.includes(query.trim()));

  if (view === "edit") {
    return (
      <section className="ban-page silence-page">
        <div className="silence-card">
          <section className="ban-form-block">
            <div className="ban-row"><span className="ban-label">禁赛类型：</span>{silenceTypes.map((item) => <label className="ban-radio" key={item}><input checked={muteTypes.includes(item)} onChange={() => toggleType(item, muteTypes, setMuteTypes)} type="checkbox" />{item}</label>)}</div>
            <label className="ban-days-row"><span className="ban-label">禁赛到：</span><input type="datetime-local" value={until} onChange={(event) => setUntil(event.target.value)} /><span>北京时间：{formatBeijingTime(until)}</span></label>
            <label className="ban-textarea-label"><span>用户ID：</span><textarea value={muteUserIds} onChange={(event) => setMuteUserIds(event.target.value)} placeholder="请输入用户ID，一行一个用户ID" /></label>
            <div className="ban-actions"><button onClick={() => addRecords("mute")} type="button">禁赛</button><button onClick={() => setMuteUserIds("")} type="button">取消</button></div>
          </section>
          <section className="ban-form-block">
            <div className="ban-row"><span className="ban-label">解禁类型：</span>{silenceTypes.map((item) => <label className="ban-radio" key={item}><input checked={unmuteTypes.includes(item)} onChange={() => toggleType(item, unmuteTypes, setUnmuteTypes)} type="checkbox" />{item}</label>)}</div>
            <label className="ban-textarea-label ban-unban-area"><span>用户ID：</span><textarea value={unmuteUserIds} onChange={(event) => setUnmuteUserIds(event.target.value)} placeholder="请输入用户ID，一行一个用户ID" /></label>
            <div className="ban-actions"><button onClick={() => addRecords("unmute")} type="button">解禁</button><button onClick={() => setUnmuteUserIds("")} type="button">取消</button></div>
          </section>
        </div>
      </section>
    );
  }

  return (
    <section className="silence-list-page">
      <button className="mail-primary-button" onClick={() => setView("edit")} type="button">禁言/解禁</button>
      <div className="silence-list-card">
        <div className="silence-search-row"><input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="用户ID" /><button type="button">搜索</button></div>
        <MailDataTable columns={["创建时间", "用户ID", "禁言至", "状态", "创建时间", "操作者"]} rows={filtered.map((record) => ({ 创建时间: record.createdAt, 用户ID: record.userId, 禁言至: record.until, 状态: record.status, 操作者: record.operator }))} />
      </div>
    </section>
  );
}

function OpenServerPage({ postWithToken }: { postWithToken: (endpoint: string, body: unknown) => Promise<ApiPostResponse> }) {
  const [svrIdQuery, setSvrIdQuery] = React.useState("");
  const [nextTime, setNextTime] = React.useState("");
  const [resetSvrId, setResetSvrId] = React.useState("");
  const [resetTime, setResetTime] = React.useState("");
  const [clearSvrId, setClearSvrId] = React.useState("");
  const [status, setStatus] = React.useState("");

  const call = async (endpoint: string, body: unknown, label: string) => {
    const result = await postWithToken(endpoint, body);
    const payload = getObject(result.payload);
    setStatus(result.ok ? `${label}已提交：${String(payload?.result ?? "请去飞书群里面查看消息")}` : `${label}失败：HTTP ${result.status}`);
  };

  return (
    <section className="open-server-page">
      <header className="open-server-hero">
        <span>admin</span>
        <h2>GM 开服管理工具</h2>
        <p>按功能卡片执行开服时间查询、下一次开服时间配置，以及需要谨慎确认的重置和清除操作。</p>
      </header>
      <OpenServerCard color="blue" title="查询开服时间" desc="输入服务器 ID 后查询当前开服时间，结果会同步发送到飞书群。">
        <input value={svrIdQuery} onChange={(event) => setSvrIdQuery(event.target.value)} placeholder="请输入服务器ID（SvrId）进行查询" />
        <button onClick={() => void call("/gmGetSvrOpenTime", { SvrId: Number(svrIdQuery) }, "查询开服时间")} type="button">查询</button>
      </OpenServerCard>
      <OpenServerCard color="green" title="获取下一次游戏开服时间" desc="获取全局下一次开服时间，并将查询结果通知到飞书群。">
        <button onClick={() => void call("/gmGetLsOpenNextTime", {}, "获取下一次开服时间")} type="button">获取下一次开服时间并通知飞书</button>
      </OpenServerCard>
      <OpenServerCard color="purple" title="设置下一次游戏开服时间" desc="选择新的下一次开服时间，确认后会写入配置并发送飞书通知。">
        <input type="datetime-local" value={nextTime} onChange={(event) => setNextTime(event.target.value)} />
        <button onClick={() => void call("/gmSetLsOpenNextTime", { time_next: Math.floor(new Date(nextTime).getTime() / 1000) }, "设置下一次开服时间")} type="button">设置并通知飞书</button>
      </OpenServerCard>
      <OpenServerCard color="red" title="重置开服时间" desc="按服务器 ID 重置开服时间，可指定重置后的开服时间；不选时间则设为 0，等新建角色时再设置开服时间。此操作需要二次确认并会发送飞书通知。">
        <input value={resetSvrId} onChange={(event) => setResetSvrId(event.target.value)} placeholder="请输入服务器ID（SvrId）" />
        <input type="datetime-local" value={resetTime} onChange={(event) => setResetTime(event.target.value)} />
        <button onClick={() => window.confirm("确认重置开服时间？") && void call("/gmResetSvrOpenTime", { SvrId: Number(resetSvrId), TimeReset: resetTime ? Math.floor(new Date(resetTime).getTime() / 1000) : 0 }, "重置开服时间")} type="button">重置开服时间</button>
      </OpenServerCard>
      <OpenServerCard color="red" title="清除指定服玩家聊天记录" desc="清除指定服务器玩家聊天记录，此操作不可恢复，请谨慎执行。">
        <input value={clearSvrId} onChange={(event) => setClearSvrId(event.target.value)} placeholder="请输入服务器ID（SvrId） 清除聊天记录" />
        <button onClick={() => window.confirm("确认清除聊天记录？") && void call("/gmClearChatMsg", { SvrId: Number(clearSvrId) }, "清除聊天记录")} type="button">清除聊天记录</button>
      </OpenServerCard>
      {status && <div className="open-server-status">{status}</div>}
    </section>
  );
}

function OpenServerCard({ children, color, desc, title }: { children: React.ReactNode; color: "blue" | "green" | "purple" | "red"; desc: string; title: string }) {
  return <section className={`open-server-card ${color}`}><span>admin</span><h3>{title}</h3><p>{desc}</p><div>{children}</div></section>;
}

function MailSuitePage({ active, canUploadItemTable, postWithToken, session, setActive }: { active: MailSectionKey; canUploadItemTable: boolean; postWithToken: (endpoint: string, body: unknown) => Promise<ApiPostResponse>; session: Session; setActive: (section: SectionKey) => void }) {
  const [items, setItems] = React.useState<ItemOption[]>([]);
  const [serverOptions, setServerOptions] = React.useState<ServerOption[]>([]);
  const [templates, setTemplates] = React.useState<MailTemplate[]>([]);
  const [rewardTemplates, setRewardTemplates] = React.useState<RewardTemplate[]>([]);
  const [mailRows, setMailRows] = React.useState<Record<string, unknown>[]>([]);
  const [localMailRows, setLocalMailRows] = React.useState<Record<string, unknown>[]>([]);
  const [view, setView] = React.useState<"list" | "edit">("list");
  const [recordTab, setRecordTab] = React.useState<"mail" | "claim">("mail");
  const [userIdQuery, setUserIdQuery] = React.useState("");
  const [templateQuery, setTemplateQuery] = React.useState("");
  const [editingMailTemplate, setEditingMailTemplate] = React.useState<MailTemplate | undefined>();
  const [editingRewardTemplate, setEditingRewardTemplate] = React.useState<RewardTemplate | undefined>();
  const [editingMailRow, setEditingMailRow] = React.useState<Record<string, unknown> | undefined>();
  const [status, setStatus] = React.useState("");

  const refreshLocalMailData = React.useCallback(async () => {
    const [itemResponse, templateResponse, rewardResponse, serverResponse, scheduledResponse] = await Promise.all([
      fetch("/local-api/items"),
      fetch("/local-api/mail-templates"),
      fetch("/local-api/reward-templates"),
      fetch("/local-api/game-servers"),
      fetch(`/local-api/scheduled-mails?serverUrl=${encodeURIComponent(session.serverUrl)}`),
    ]);
    const itemPayload = (await itemResponse.json().catch(() => ({}))) as { items?: ItemOption[] };
    const templatePayload = (await templateResponse.json().catch(() => ({}))) as { templates?: MailTemplate[] };
    const rewardPayload = (await rewardResponse.json().catch(() => ({}))) as { templates?: RewardTemplate[] };
    const serverPayload = (await serverResponse.json().catch(() => ({}))) as { servers?: ServerOption[] };
    const scheduledPayload = (await scheduledResponse.json().catch(() => ({}))) as { mails?: Record<string, unknown>[] };
    setItems(itemPayload.items ?? []);
    setTemplates(templatePayload.templates ?? []);
    setRewardTemplates(rewardPayload.templates ?? []);
    setServerOptions(serverPayload.servers ?? []);
    setLocalMailRows(scheduledPayload.mails ?? []);
  }, [session.serverUrl]);

  const refreshMailList = React.useCallback(async () => {
    const result = await postWithToken("/gmMailLst", {});
    const error = apiBusinessError(result);
    if (error) {
      setStatus(`邮件列表读取失败：${error}`);
      setMailRows([]);
      return;
    }
    const data = getApiData(result.payload);
    const rows = getArray(data?.Lst).filter((row): row is Record<string, unknown> => Boolean(getObject(row)));
    setMailRows(rows);
  }, [postWithToken]);

  React.useEffect(() => {
    void refreshLocalMailData().catch(() => undefined);
  }, [refreshLocalMailData]);

  React.useEffect(() => {
    if (active === "mailPersonal" || active === "mailGlobal") {
      void refreshMailList().catch(() => setMailRows([]));
    }
  }, [active, refreshMailList]);

  React.useEffect(() => {
    setView("list");
    setEditingMailTemplate(undefined);
    setEditingRewardTemplate(undefined);
    setEditingMailRow(undefined);
    setStatus("");
  }, [active]);

  const uploadItemTable = async (file: File) => {
    const formData = new FormData();
    formData.append("file", file);
    const response = await fetch("/local-api/items/upload", { method: "POST", body: formData });
    const payload = (await response.json().catch(() => ({}))) as { items?: ItemOption[]; error?: string };
    if (!response.ok) throw new Error(payload.error || "上传道具表失败");
    setItems(payload.items ?? []);
    setStatus(`道具表已导入，共 ${payload.items?.length ?? 0} 个道具`);
  };

  if (active === "mailTemplate") {
    return view === "edit"
      ? <MailTemplateEditor onBack={() => setView("list")} onSaved={() => { setView("list"); setEditingMailTemplate(undefined); void refreshLocalMailData(); }} template={editingMailTemplate} />
      : <MailTemplateList query={templateQuery} setQuery={setTemplateQuery} templates={templates} onCreate={() => { setEditingMailTemplate(undefined); setView("edit"); }} onEdit={(template) => { setEditingMailTemplate(template); setView("edit"); }} onDelete={async (template) => {
        if (!window.confirm(`确认删除邮件模板「${template.name}」？`)) return;
        await fetch(`/local-api/mail-templates/${encodeURIComponent(template.id)}`, { method: "DELETE" });
        await refreshLocalMailData();
      }} />;
  }

  if (active === "mailRewardTemplate") {
    return view === "edit"
      ? <RewardTemplateEditor canUploadItemTable={canUploadItemTable} items={items} onBack={() => setView("list")} onSaved={() => { setView("list"); setEditingRewardTemplate(undefined); void refreshLocalMailData(); }} onUploadItemTable={uploadItemTable} template={editingRewardTemplate} />
      : <RewardTemplateList templates={rewardTemplates} onCreate={() => { setEditingRewardTemplate(undefined); setView("edit"); }} onEdit={(template) => { setEditingRewardTemplate(template); setView("edit"); }} onDelete={async (template) => {
        if (!window.confirm(`确认删除奖励模板「${template.title}」？`)) return;
        await fetch(`/local-api/reward-templates/${encodeURIComponent(template.id)}`, { method: "DELETE" });
        await refreshLocalMailData();
      }} />;
  }

  if (view === "edit") {
    return (
      <MailEditor
        global={active === "mailGlobal"}
        canUploadItemTable={canUploadItemTable}
        items={items}
        serverOptions={serverOptions}
        rewardTemplates={rewardTemplates}
        templates={templates}
        initialMail={editingMailRow}
        onBack={() => setView("list")}
        onSubmit={async (body) => {
          const submitted = getObject(body) ?? {};
          const startSeconds = Number(submitted.St);
          const isScheduled = Number.isFinite(startSeconds) && startSeconds > Math.floor(Date.now() / 1000) + 5;
          if (isScheduled) {
            const response = await fetch("/local-api/scheduled-mails", {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({
                serverUrl: session.serverUrl,
                game: session.game,
                serverName: session.serverName,
                operator: session.operatorAccount,
                scheduleAt: startSeconds,
                body,
              }),
            });
            const payload = (await response.json().catch(() => ({}))) as { mail?: Record<string, unknown>; error?: string };
            if (!response.ok || !payload.mail) throw new Error(payload.error || "创建定时邮件失败");
            const localRow = payload.mail;
            const message = `邮件已加入定时队列，将在 ${formatTimestamp(startSeconds)} 发送`;
            setStatus(message);
            setView("list");
            setEditingMailRow(undefined);
            setLocalMailRows((current) => [localRow, ...current.filter((row) => String(row.Id) !== String(localRow.Id))].slice(0, 50));
            return message;
          }
          const result = await postWithToken("/gmMailAdd", body);
          const error = apiBusinessError(result);
          if (error) {
            throw new Error(`邮件提交失败：${error}`);
          }
          const data = getApiData(result.payload);
          const localRow = {
            ...submitted,
            ...(data ?? {}),
            Id: String(data?.Id ?? submitted.Id ?? `local-${Date.now()}`),
            RegtEnd: data?.RegtEnd ?? submitted.Regt,
            CreateTime: data?.CreateTime ?? Math.floor(Date.now() / 1000),
            __local: !data?.Id,
            __claimed: false,
          };
          const message = `邮件已提交到服务器${data?.Id ? `，ID：${String(data.Id)}` : ""}`;
          setStatus(message);
          const submittedTyp = Number(getObject(body)?.Typ);
          const submittedTargets = getArray(getObject(body)?.TargetID);
          setView("list");
          setEditingMailRow(undefined);
          setLocalMailRows((current) => [localRow, ...current.filter((row) => String(row.Id) !== String(localRow.Id))].slice(0, 20));
          await refreshMailList();
          if ((submittedTyp === 2 || (submittedTyp === 1 && submittedTargets.length === 0)) && active !== "mailGlobal") {
            setActive("mailGlobal");
          } else if (submittedTyp === 3 && submittedTargets.length > 0 && active !== "mailPersonal") {
            setActive("mailPersonal");
          }
          setStatus(message);
          return message;
        }}
        onUploadItemTable={uploadItemTable}
      />
    );
  }

  return (
    <MailListPage
      active={active}
      serverOptions={serverOptions}
      mailRows={mailRows}
      localMailRows={localMailRows}
      recordTab={recordTab}
      setRecordTab={setRecordTab}
      status={status}
      userIdQuery={userIdQuery}
      setUserIdQuery={setUserIdQuery}
      onCreate={() => { setEditingMailRow(undefined); setView("edit"); }}
      onEdit={(row) => { setEditingMailRow(row); setView("edit"); }}
      onDelete={async (id) => {
        const localRow = localMailRows.find((row) => String(row.Id) === id);
        if (localRow?.__scheduled) {
          const response = await fetch(`/local-api/scheduled-mails/${encodeURIComponent(id)}`, { method: "DELETE" });
          const payload = (await response.json().catch(() => ({}))) as { error?: string };
          if (!response.ok) {
            setStatus(`撤回失败：${payload.error || "定时邮件删除失败"}`);
            return;
          }
          setLocalMailRows((current) => current.filter((row) => String(row.Id) !== id));
          setStatus("定时邮件已撤回");
          return;
        }
        const result = await postWithToken("/gmMailDel", { Id: id });
        const error = apiBusinessError(result);
        setStatus(error ? `删除失败：${error}` : "邮件已删除");
        if (!error) setLocalMailRows((current) => current.filter((row) => String(row.Id) !== id));
        await refreshMailList();
      }}
      onRefresh={() => void refreshMailList()}
    />
  );
}

function MailListPage({ active, localMailRows, mailRows, onCreate, onDelete, onEdit, onRefresh, recordTab, serverOptions, setRecordTab, setUserIdQuery, status, userIdQuery }: { active: MailSectionKey; localMailRows: Record<string, unknown>[]; mailRows: Record<string, unknown>[]; onCreate: () => void; onDelete: (id: string) => Promise<void>; onEdit: (row: Record<string, unknown>) => void; onRefresh: () => void; recordTab: "mail" | "claim"; serverOptions: ServerOption[]; setRecordTab: (tab: "mail" | "claim") => void; setUserIdQuery: (value: string) => void; status: string; userIdQuery: string }) {
  const global = active === "mailGlobal";
  const mergedRows = [...localMailRows, ...mailRows.filter((row) => !localMailRows.some((localRow) => String(localRow.Id) === String(row.Id)))];
  const rows = mergedRows.filter((row) => {
    const targetIds = getArray(row.TargetID);
    const typ = Number(row.Typ);
    const isServerMail = typ === 2;
    const isPersonalMail = typ === 3 || (!typ && targetIds.length > 0);
    if (global && isPersonalMail) return false;
    if (!global && !isPersonalMail) return false;
    if (!userIdQuery.trim()) return true;
    return formatCell(row.TargetID).includes(userIdQuery.trim());
  });
  const listColumns = ["ID", "模板名称", "目标", "状态", "时间", "操作"];

  if (global && recordTab === "claim") {
    return (
      <section className="mail-page">
        <div className="mail-tabs"><button onClick={() => setRecordTab("mail")} type="button">全局邮件</button><button className="active" type="button">领取记录</button></div>
        <div className="mail-filter-line"><label>用户 ID：<input value={userIdQuery} onChange={(event) => setUserIdQuery(event.target.value)} /></label><button onClick={onRefresh} type="button"><Search size={14} />Search</button></div>
        <MailDataTable columns={["邮件ID", "模板名称", "是否查看", "是否领取", "查看时间", "领取时间", "生效时间", "过期时间"]} rows={[]} />
      </section>
    );
  }

  return (
    <section className="mail-page">
      {global && <div className="mail-tabs"><button className="active" type="button">全局邮件</button><button onClick={() => setRecordTab("claim")} type="button">领取记录</button></div>}
      {!global && <div className="mail-filter-line"><label>用户 ID：<input value={userIdQuery} onChange={(event) => setUserIdQuery(event.target.value)} /></label><button onClick={onRefresh} type="button"><Search size={14} />Search</button></div>}
      <section className="mail-table-card">
        <button className="mail-primary-button" onClick={onCreate} type="button">新建</button>
        <MailDataTable
          columns={listColumns}
          rows={rows.map((row) => {
            const id = String(row.Id ?? row.id ?? "");
            const targetIds = getArray(row.TargetID);
            const typ = Number(row.Typ);
            const isServerMail = typ === 2;
            const typeName = isServerMail ? "区服" : typ === 3 || (!typ && targetIds.length > 0) ? "个人" : "全局";
            const title = formatCell(row.Titel ?? row.Title ?? "自定义邮件");
            const state = Number(row.St);
            const statusText = row.__scheduled
              ? row.__scheduledStatus === "failed" ? `定时失败：${formatCell(row.__scheduledError)}` : "待发送"
              : row.__claimed ? "已领取" : state > 0 ? `状态 ${state}` : "未领取";
            const createdAt = formatCell(row.CreateTime ?? row.CreatedAt ?? row.Ct ?? row.createdAt);
            const regEnd = row.RegtEnd ?? row.Regt;
            return {
              ID: <span className="mail-id-cell"><span>{id}</span><em>{typeName}</em></span>,
              模板名称: title,
              目标: <MailTargetSummary row={row} serverOptions={serverOptions} />,
              状态: statusText,
              时间: <span className="mail-time-cell"><span>注册开始: {formatTimestamp(row.RegtBegin)}</span><span>注册结束: {formatTimestamp(regEnd)}</span><span>生效时间: {formatTimestamp(row.St)}</span><span>过期时间: {formatTimestamp(row.Et)}</span><span>创建时间: {formatTimestampValue(createdAt)}</span></span>,
              操作: <div className="mail-action-buttons"><button onClick={() => onEdit(row)} type="button">编辑</button><button onClick={() => void onDelete(id)} type="button">撤回</button></div>,
            };
          })}
        />
      </section>
      {status && <div className="mail-status">{status}</div>}
    </section>
  );
}

function MailDataTable({ columns, rows }: { columns: string[]; rows: Array<Record<string, React.ReactNode>> }) {
  return (
    <table className="mail-data-table">
      <thead><tr>{columns.map((column) => <th key={column}>{column}</th>)}</tr></thead>
      <tbody>{rows.length ? rows.map((row, index) => <tr key={index}>{columns.map((column) => <td key={column}>{row[column] ?? "暂无数据"}</td>)}</tr>) : <tr><td colSpan={columns.length}>暂无数据</td></tr>}</tbody>
    </table>
  );
}

function MailTargetSummary({ row, serverOptions }: { row: Record<string, unknown>; serverOptions: ServerOption[] }) {
  const targetIds = getArray(row.TargetID);
  const typ = Number(row.Typ);
  const isServerMail = typ === 2;
  const isPersonalMail = typ === 3 || (!typ && targetIds.length > 0);
  const targetText = targetIds.length
    ? isServerMail
      ? targetIds.map((targetId) => serverOptions.find((server) => server.id === Number(targetId))?.name ?? `游戏内区服 ${String(targetId)}`).join(", ")
      : `用户 ${formatCell(row.TargetID)}`
    : isPersonalMail
      ? "未填写用户"
      : "全服";
  const platforms = getArray(row.Platform).map((item) => Number(item)).filter((item) => item === 1 || item === 2);
  const versions = getArray(row.Version).map((item) => String(item)).filter(Boolean);
  const regBegin = Number(row.RegtBegin);
  const regEnd = Number(row.RegtEnd ?? row.Regt);
  const tags = [
    { label: "目标", value: targetText },
    platforms.length ? { label: "系统", value: platforms.map((item) => item === 1 ? "GooglePlay" : "iOS").join(", ") } : null,
    versions.length ? { label: "版本", value: versions.join(", ") } : null,
    regBegin > 0 ? { label: "注册开始", value: formatTimestamp(regBegin) } : null,
    regEnd > 0 ? { label: "注册结束", value: formatTimestamp(regEnd) } : null,
  ].filter((item): item is { label: string; value: string } => Boolean(item));
  return <span className="mail-target-summary">{tags.map((item) => <span key={item.label}><em>{item.label}</em>{item.value}</span>)}</span>;
}

function MailEditor({ canUploadItemTable, global, initialMail, items, onBack, onSubmit, onUploadItemTable, rewardTemplates, serverOptions, templates }: { canUploadItemTable: boolean; global: boolean; initialMail?: Record<string, unknown>; items: ItemOption[]; onBack: () => void; onSubmit: (body: unknown) => Promise<string | void>; onUploadItemTable: (file: File) => Promise<void>; rewardTemplates: RewardTemplate[]; serverOptions: ServerOption[]; templates: MailTemplate[] }) {
  const now = new Date();
  const [mailType, setMailType] = React.useState(global ? "global" : "personal");
  const [templateId, setTemplateId] = React.useState("custom");
  const [rewardTemplateId, setRewardTemplateId] = React.useState("custom");
  const [title, setTitle] = React.useState(String(initialMail?.Titel ?? initialMail?.Title ?? ""));
  const [body, setBody] = React.useState(String(initialMail?.Body ?? ""));
  const [targetIds, setTargetIds] = React.useState(getArray(initialMail?.TargetID).join(","));
  const [sendMode, setSendMode] = React.useState<"now" | "scheduled">(initialMail?.__scheduled ? "scheduled" : "now");
  const [rewards, setRewards] = React.useState<MailRewardItem[]>(() => {
    const raw = getArray(initialMail?.ItemLst);
    const parsed: MailRewardItem[] = [];
    for (let index = 0; index < raw.length; index += 2) {
      const itemId = Number(raw[index]);
      const count = Number(raw[index + 1]);
      if (Number.isFinite(itemId) && Number.isFinite(count) && itemId > 0 && count > 0) parsed.push({ itemId: String(itemId), count: String(count) });
    }
    return parsed.length ? parsed : [{ itemId: "", count: "0" }];
  });
  const [startTime, setStartTime] = React.useState(secondsToDatetimeLocal(initialMail?.St) || toDatetimeLocal(now));
  const [endTime, setEndTime] = React.useState(secondsToDatetimeLocal(initialMail?.Et) || MAIL_DEFAULT_EXPIRE);
  const [filterRows, setFilterRows] = React.useState<Array<{ id: number; field: string; op: string; value: string }>>([]);
  const [error, setError] = React.useState("");
  const [submitting, setSubmitting] = React.useState(false);
  const isGlobalMail = mailType === "global";
  const defaultFilterValue = (field: string, op = "=") => {
    if (field === "regTime") return (op === "<=" ? defaultRegEndTime() : MAIL_DEFAULT_REG_BEGIN).slice(0, 10);
    return "";
  };
  const filterFieldOptions = [
    { value: "system", label: "系统" },
    { value: "version", label: "app版本" },
    { value: "regTime", label: "注册时间" },
    { value: "server", label: "游戏内区服", disabled: !isGlobalMail },
    { value: "language", label: "app语言", disabled: true },
    { value: "country", label: "国家", disabled: true },
  ];

  const selectedTemplate = templates.find((template) => template.id === templateId);
  const selectedRewardTemplate = rewardTemplates.find((template) => template.id === rewardTemplateId);
  const selectedTemplateContents = selectedTemplate?.contents;

  React.useEffect(() => {
    if (templateId !== "custom" && selectedTemplate) {
      const content = templatePrimaryContent(selectedTemplate);
      setTitle(content.title ?? selectedTemplate.title);
      setBody(content.body ?? selectedTemplate.body);
    }
  }, [selectedTemplate, templateId]);

  React.useEffect(() => {
    if (rewardTemplateId !== "custom" && selectedRewardTemplate) {
      setRewards(selectedRewardTemplate.items.length ? selectedRewardTemplate.items : [{ itemId: "", count: "0" }]);
    }
  }, [rewardTemplateId, selectedRewardTemplate]);

  const submit = async () => {
    const targets = isGlobalMail ? [] : toNumberArray(targetIds);
    if (!isGlobalMail && !targets.length) {
      setError("请输入用户 ID");
      return;
    }
    if (!title.trim()) {
      setError("请输入邮件标题");
      return;
    }
    const endSeconds = parseDatetimeLocalSeconds(endTime);
    if (!endSeconds) {
      setError("请选择有效的过期时间");
      return;
    }
    const rewardValidation = validateRewardRows(rewards, items);
    if (!rewardValidation.ok) {
      setError(rewardValidation.message ?? "请填写有效的奖励道具和数量，或选择无奖励");
      return;
    }
    const conditionRows = filterRows.filter((row) => isGlobalMail || row.field !== "server");
    const versionList = conditionRows.filter((row) => row.field === "version").flatMap((row) => toVersionNumberArray(row.value));
    const platformList = conditionRows.filter((row) => row.field === "system").flatMap((row) => toPlatformNumberArray(row.value));
    const serverTargetIds = isGlobalMail ? filterRows.filter((row) => row.field === "server").flatMap((row) => toFlexibleNumberArray(row.value)) : [];
    const regBeginValues = conditionRows.filter((row) => row.field === "regTime" && row.op === ">=" && row.value).map((row) => parseDatetimeLocalSeconds(dateToDatetimeLocal(row.value)));
    const regEndValues = conditionRows.filter((row) => row.field === "regTime" && row.op === "<=" && row.value).map((row) => parseDatetimeLocalSeconds(dateToDatetimeLocal(row.value, true)));
    const regBeginSeconds = regBeginValues[0] ?? parseDatetimeLocalSeconds(MAIL_DEFAULT_REG_BEGIN);
    const regEndSeconds = regEndValues[0] ?? parseDatetimeLocalSeconds(defaultRegEndTime());
    if (regBeginValues.some((value) => !value) || regEndValues.some((value) => !value)) {
      setError("请选择有效的注册时间区间");
      return;
    }
    if (regBeginSeconds && regEndSeconds && regEndSeconds <= regBeginSeconds) {
      setError("注册结束时间必须晚于注册开始时间");
      return;
    }
    if (conditionRows.some((row) => row.field !== "regTime" && row.op !== "=" && row.value.trim())) {
      setError("当前邮件接口只支持系统、版本、游戏内区服的等于条件，范围和排除条件暂不支持保存");
      return;
    }
    const emptyCondition = conditionRows.find((row) => ["system", "version", "server"].includes(row.field) && !row.value.trim());
    if (emptyCondition) {
      const labels: Record<string, string> = { system: "系统", version: "版本", server: "区服" };
      setError(`${labels[emptyCondition.field] ?? "条件"}未填写`);
      return;
    }
    if (conditionRows.some((row) => row.field === "version" && row.value.trim() && !toVersionNumberArray(row.value).length)) {
      setError("APP版本请填写 x.x.x 或 x.x.x.x 格式，例如 1.8.0.0");
      return;
    }
    if (conditionRows.some((row) => row.field === "system" && row.value.trim() && !toPlatformNumberArray(row.value).length)) {
      setError("系统请选择 GooglePlay 或 iOS");
      return;
    }
    if (isGlobalMail && filterRows.some((row) => row.field === "server" && row.value.trim()) && !serverTargetIds.length) {
      setError("区服未填写");
      return;
    }
    if (isGlobalMail && serverTargetIds.length && serverOptions.length) {
      const serverIds = new Set(serverOptions.map((server) => Number(server.id)));
      const missingServer = serverTargetIds.find((serverId) => !serverIds.has(Number(serverId)));
      if (missingServer) {
        setError(`区服不存在：${missingServer}`);
        return;
      }
    }
    if (isGlobalMail && filterRows.some((row) => row.field === "server" && !row.value.trim())) {
      setError("区服未填写");
      return;
    }
    const startSeconds = sendMode === "scheduled" ? parseDatetimeLocalSeconds(startTime) : 0;
    if (sendMode === "scheduled" && !startSeconds) {
      setError("请选择有效的定时发送时间");
      return;
    }
    if (sendMode === "scheduled" && startSeconds <= Math.floor(Date.now() / 1000) + 5) {
      setError("定时发送时间必须晚于当前时间");
      return;
    }
    if (sendMode === "scheduled" && endSeconds <= startSeconds) {
      setError("过期时间必须晚于生效时间");
      return;
    }
    const mailTyp = isGlobalMail ? serverTargetIds.length ? 2 : 1 : 3;
    const mailTargets = isGlobalMail ? serverTargetIds : targets;
    setError("");
    setSubmitting(true);
    try {
      await onSubmit({
        Typ: mailTyp,
        TargetID: mailTargets,
        RegtBegin: regBeginSeconds,
        Regt: regEndSeconds,
        Et: endSeconds,
        St: sendMode === "scheduled" ? startSeconds : 0,
        SenderName: "GM",
        Titel: title,
        Body: body,
        BodyData: [],
        BodyData2: selectedTemplateContents ? [JSON.stringify({ type: "multiLanguage", contents: selectedTemplateContents })] : [],
        ItemLst: rewardValidation.itemList,
        Platform: platformList,
        Version: versionList,
      });
      setError("保存成功，正在返回列表...");
    } catch (submitError) {
      setError(submitError instanceof Error ? submitError.message : "保存失败");
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <section className="mail-edit-page">
      <div className="mail-edit-card">
        <header>Create Mail</header>
        <div className="mail-form">
          <div className="mail-form-row"><span>邮件类型</span><input disabled value={isGlobalMail ? "全局邮件" : "个人邮件"} /></div>
          <label className="mail-form-row"><span>邮件模板</span><select value={templateId} onChange={(event) => setTemplateId(event.target.value)}><option value="custom">自定义</option>{templates.map((template) => <option key={template.id} value={template.id}>{template.name}</option>)}</select></label>
          <div className="mail-form-row mail-filter-builder">
            <span>条件</span>
            <div className="mail-condition-list">
              {filterRows.length === 0 && <div className="mail-condition-empty">默认无条件，邮件会发给当前类型下的全部目标。</div>}
              {filterRows.map((row) => {
                const unsupported = row.field === "language" || row.field === "country";
                const comparisonOps = row.field === "version" || row.field === "server";
                return (
                  <div className="mail-condition-row" key={row.id}>
                    <select value={row.field} onChange={(event) => {
                      const nextField = !isGlobalMail && event.target.value === "server" ? "version" : event.target.value;
                      const nextOp = nextField === "regTime" ? ">=" : "=";
                      setFilterRows((current) => current.map((item) => item.id === row.id ? { ...item, field: nextField, op: nextOp, value: defaultFilterValue(nextField, nextOp) } : item));
                    }}>
                      {filterFieldOptions.map((option) => <option disabled={option.disabled} key={option.value} value={option.value}>{option.label}</option>)}
                    </select>
                    {row.field === "regTime" || comparisonOps ? (
                      <select className="mail-condition-expression" value={row.op} onChange={(event) => {
                        const nextOp = event.target.value;
                        setFilterRows((current) => current.map((item) => item.id === row.id ? { ...item, op: nextOp, value: item.value || defaultFilterValue(item.field, nextOp) } : item));
                      }}>
                        {comparisonOps && <option value="=">=</option>}
                        <option value=">=">&gt;=</option>
                        <option value="<=">&lt;=</option>
                        {comparisonOps && <option value="!=">!=</option>}
                      </select>
                    ) : (
                      <div className="mail-condition-op" role="group" aria-label="条件操作">
                        <button className={row.op === "=" ? "active" : ""} onClick={() => setFilterRows((current) => current.map((item) => item.id === row.id ? { ...item, op: "=" } : item))} type="button">=</button>
                        <button className={row.op === "!=" ? "active" : ""} onClick={() => setFilterRows((current) => current.map((item) => item.id === row.id ? { ...item, op: "!=" } : item))} type="button">排除</button>
                      </div>
                    )}
                    {row.field === "system" ? (
                      <select value={row.value} onChange={(event) => setFilterRows((current) => current.map((item) => item.id === row.id ? { ...item, value: event.target.value } : item))}>
                        <option value="">请选择</option>
                        <option value="1">GooglePlay</option>
                        <option value="2">iOS</option>
                      </select>
                    ) : row.field === "regTime" ? (
                      <input type="date" value={row.value ? row.value.slice(0, 10) : ""} onChange={(event) => setFilterRows((current) => current.map((item) => item.id === row.id ? { ...item, value: event.target.value } : item))} />
                    ) : row.field === "server" && serverOptions.length ? (
                      <select value={row.value} onChange={(event) => setFilterRows((current) => current.map((item) => item.id === row.id ? { ...item, value: event.target.value } : item))}>
                        <option value="">请选择游戏内区服</option>
                        {serverOptions.map((server) => <option key={server.id} value={server.id}>{server.name}</option>)}
                      </select>
                    ) : (
                      <input disabled={unsupported} value={row.value} onChange={(event) => setFilterRows((current) => current.map((item) => item.id === row.id ? { ...item, value: event.target.value } : item))} placeholder={unsupported ? "当前接口暂未开放" : row.field === "version" ? "例如 1.8.0.0" : row.field === "server" ? "例如 12 或 1,2" : "例如 1,2"} />
                    )}
                    <button className="mail-condition-remove" onClick={() => setFilterRows((current) => current.filter((item) => item.id !== row.id))} type="button">删除</button>
                  </div>
                );
              })}
              <button className="mail-add-condition" onClick={() => setFilterRows((current) => [...current, { id: Date.now() + current.length, field: "version", op: "=", value: "" }])} type="button">新增条件</button>
              <small className="mail-condition-hint">{isGlobalMail ? "多个条件同时填写时为且的关系。填写“游戏内区服”后，会按接口 Typ=2 将 TargetID 作为区服ID发送。" : "多个条件同时填写时为且的关系。个人邮件不会提交游戏内区服条件。"}</small>
            </div>
          </div>
          <label className="mail-form-row"><span>奖励模板</span><select value={rewardTemplateId} onChange={(event) => setRewardTemplateId(event.target.value)}><option value="custom">自定义</option>{rewardTemplates.map((template) => <option key={template.id} value={template.id}>{template.title}</option>)}</select></label>
          <RewardRows canUploadItemTable={canUploadItemTable} items={items} onUploadItemTable={onUploadItemTable} rewards={rewards} setRewards={setRewards} />
          {!isGlobalMail && <label className="mail-form-row mail-textarea-row"><span>用户 ID</span><textarea value={targetIds} onChange={(event) => setTargetIds(event.target.value)} /></label>}
          <label className="mail-form-row"><span>邮件标题</span><input value={title} onChange={(event) => setTitle(event.target.value)} /></label>
          <label className="mail-form-row mail-textarea-row"><span>邮件内容</span><textarea value={body} onChange={(event) => setBody(event.target.value)} /></label>
          <label className="mail-form-row"><span>发送方式</span><select value={sendMode} onChange={(event) => setSendMode(event.target.value === "scheduled" ? "scheduled" : "now")}><option value="now">立即发送</option><option value="scheduled">定时发送</option></select></label>
          {sendMode === "scheduled" && <label className="mail-form-row"><span>定时发送</span><input type="datetime-local" value={startTime} onChange={(event) => setStartTime(event.target.value)} /><em>北京时间 {formatBeijingTime(startTime)}</em></label>}
          <label className="mail-form-row"><span>过期时间</span><input type="datetime-local" value={endTime} onChange={(event) => setEndTime(event.target.value)} /><em>北京时间 {formatBeijingTime(endTime)}，{formatTimeDistance(endTime)}</em></label>
          {error && <div className="mail-form-error">{error}</div>}
          <div className="mail-form-actions"><button disabled={submitting} onClick={() => void submit()} type="button">{submitting ? "发送中..." : "发送"}</button><button disabled={submitting} onClick={onBack} type="button">取消</button></div>
        </div>
      </div>
    </section>
  );
}

function RewardRows({ canUploadItemTable = true, items, onUploadItemTable, rewards, setRewards }: { canUploadItemTable?: boolean; items: ItemOption[]; onUploadItemTable: (file: File) => Promise<void>; rewards: MailRewardItem[]; setRewards: (items: MailRewardItem[]) => void }) {
  const [uploading, setUploading] = React.useState(false);
  const [uploadError, setUploadError] = React.useState("");
  const updateReward = (index: number, patch: Partial<MailRewardItem>) => {
    setRewards(rewards.map((reward, rewardIndex) => rewardIndex === index ? { ...reward, ...patch } : reward));
  };
  const deleteReward = (index: number) => {
    setRewards(rewards.length > 1 ? rewards.filter((_, rewardIndex) => rewardIndex !== index) : [{ itemId: "", count: "0" }]);
  };
  return (
    <div className="mail-reward-block">
      {rewards.map((reward, index) => (
        <div className="mail-form-row mail-reward-row" key={index}>
          <span>{index === 0 ? "奖励内容" : ""}</span>
          <label>Item<ItemInput items={items} value={reward.itemId} onChange={(value) => updateReward(index, { itemId: value })} /></label>
          <label>Count<input inputMode="numeric" max={MAX_REWARD_COUNT} min={1} type="number" value={reward.count} onChange={(event) => updateReward(index, { count: event.target.value })} /></label>
          <button className="mail-delete-reward" onClick={() => deleteReward(index)} title="删除奖励" type="button">删除</button>
        </div>
      ))}
      <div className="mail-form-row mail-add-reward-row"><span /><button className="mail-add-reward" onClick={() => setRewards([...rewards, { itemId: "", count: "0" }])} type="button">新增</button></div>
      {canUploadItemTable && <div className="mail-form-row"><span /><label className={`mail-upload-link ${uploading ? "disabled" : ""}`}>{uploading ? "上传中..." : "上传Item表"}<input accept=".xlsx,.xls" disabled={uploading} onChange={(event) => {
        const file = event.target.files?.[0];
        event.currentTarget.value = "";
        if (!file) return;
        setUploading(true);
        setUploadError("");
        void onUploadItemTable(file).catch((error) => setUploadError(error instanceof Error ? error.message : "上传Item表失败")).finally(() => setUploading(false));
      }} type="file" /></label><small>{uploadError || (items.length ? `已加载 ${items.length} 个道具` : "未上传道具表")}</small></div>
      }
    </div>
  );
}

function ItemInput({ items, onChange, value }: { items: ItemOption[]; onChange: (value: string) => void; value: string }) {
  const listId = React.useId();
  const normalize = (raw: string) => {
    const match = raw.match(/^\s*(\d+)/);
    return match ? match[1] : raw;
  };
  const selected = items.find((item) => String(item.id) === value);
  const [draft, setDraft] = React.useState("");
  const [focused, setFocused] = React.useState(false);
  React.useEffect(() => {
    if (!focused) setDraft(selected ? `${selected.id} - ${selected.name || "未命名道具"}` : value);
  }, [focused, selected, value]);
  return (
    <div className="mail-item-picker">
      <input
        list={listId}
        onBlur={(event) => {
          setFocused(false);
          onChange(normalize(event.target.value));
        }}
        onChange={(event) => {
          const next = event.target.value;
          setDraft(next);
          onChange(next.trim() ? normalize(next) : "");
        }}
        onFocus={() => {
          setFocused(true);
          setDraft(value);
        }}
        placeholder="输入道具ID或名称"
        value={focused ? draft : selected ? `${selected.id} - ${selected.name || "未命名道具"}` : value}
      />
      <datalist id={listId}>
        {items.map((item) => <option key={item.id} value={`${item.id} - ${item.name || "未命名道具"}`} />)}
      </datalist>
      <small className={selected ? "matched" : ""}>{value ? selected ? "已匹配" : "未匹配道具" : "可手填或选择"}</small>
    </div>
  );
}

function MailTemplateList({ onCreate, onDelete, onEdit, query, setQuery, templates }: { onCreate: () => void; onDelete: (template: MailTemplate) => Promise<void>; onEdit: (template: MailTemplate) => void; query: string; setQuery: (value: string) => void; templates: MailTemplate[] }) {
  const rows = templates.filter((template) => !query.trim() || template.name.includes(query.trim()));
  return <section className="mail-page"><div className="mail-filter-line"><label>模板名称：<input value={query} onChange={(event) => setQuery(event.target.value)} /></label><button type="button"><Search size={14} />Search</button></div><section className="mail-table-card"><button className="mail-primary-button" onClick={onCreate} type="button">新建</button><MailDataTable columns={["ID", "名称", "标题", "创建时间", "更新时间", "操作"]} rows={rows.map((template) => ({ ID: template.id, 名称: <span className="mail-ellipsis-cell" title={template.name}>{template.name}</span>, 标题: <span className="mail-ellipsis-cell" title={templatePrimaryContent(template).title}>{templatePrimaryContent(template).title || "暂无数据"}</span>, 创建时间: formatTimestampValue(template.createdAt), 更新时间: formatTimestampValue(template.updatedAt), 操作: <div className="mail-action-buttons"><button onClick={() => onEdit(template)} type="button">编辑</button><button onClick={() => void onDelete(template)} type="button">删除</button></div> }))} /></section></section>;
}

function MailTemplateEditor({ onBack, onSaved, template }: { onBack: () => void; onSaved: () => void; template?: MailTemplate }) {
  const [name, setName] = React.useState(template?.name ?? "");
  const [error, setError] = React.useState("");
  const [saving, setSaving] = React.useState(false);
  const [activeLanguage, setActiveLanguage] = React.useState(defaultMailLanguage);
  const [contents, setContents] = React.useState<Record<string, MailTemplateContent>>(() => {
    const initial = Object.fromEntries(mailLanguages.map((language) => [language, { title: "", body: "" }])) as Record<string, MailTemplateContent>;
    if (template?.contents) {
      for (const language of mailLanguages) initial[language] = { title: template.contents[language]?.title ?? "", body: template.contents[language]?.body ?? "" };
    }
    if (!initial[defaultMailLanguage].title && !initial[defaultMailLanguage].body && template) {
      initial[defaultMailLanguage] = { title: template.title ?? "", body: template.body ?? "" };
    }
    return initial;
  });
  const activeContent = contents[activeLanguage] ?? { title: "", body: "" };
  const updateContent = (patch: Partial<MailTemplateContent>) => {
    setContents((current) => ({ ...current, [activeLanguage]: { ...(current[activeLanguage] ?? { title: "", body: "" }), ...patch } }));
  };
  const downloadTemplate = () => {
    const workbook = XLSX.utils.book_new();
    const header = ["模板名称", ...mailLanguages.flatMap((language) => [`${language}标题`, `${language}内容`])];
    const example = ["示例模板", ...mailLanguages.flatMap((language) => [`${language}邮件标题`, `${language}邮件内容`])];
    const sheet = XLSX.utils.aoa_to_sheet([header, example]);
    XLSX.utils.book_append_sheet(workbook, sheet, "邮件模板");
    XLSX.writeFile(workbook, "邮件模板导入模板.xlsx");
  };
  const uploadTemplate = async (file: File) => {
    try {
      const parsed = await parseMailTemplateFile(file);
      if (parsed.name && !name.trim()) setName(parsed.name);
      setContents(parsed.contents);
      setError("");
    } catch (uploadError) {
      setError(uploadError instanceof Error ? uploadError.message : "上传模板失败");
    }
  };
  const save = async () => {
    if (saving) return;
    const cleanName = name.trim();
    if (!cleanName) {
      setError("请输入模板名称");
      return;
    }
    if (cleanName.length > MAX_TEMPLATE_NAME_LENGTH) {
      setError(`模板名称最多${MAX_TEMPLATE_NAME_LENGTH}个字符`);
      return;
    }
    const cleanedContents = Object.fromEntries(mailLanguages.map((language) => [language, { title: contents[language]?.title.trim() ?? "", body: contents[language]?.body.trim() ?? "" }])) as Record<string, MailTemplateContent>;
    const hasContent = Object.values(cleanedContents).some((content) => content.title && content.body);
    if (!hasContent) {
      setError("请至少填写一个语言的邮件标题和邮件内容");
      return;
    }
    const incompleteLanguage = mailLanguages.find((language) => {
      const content = cleanedContents[language];
      return Boolean(content.title) !== Boolean(content.body);
    });
    if (incompleteLanguage) {
      setError(`${incompleteLanguage} 的邮件标题和邮件内容需要同时填写`);
      return;
    }
    setError("");
    setSaving(true);
    try {
      const primary = cleanedContents[defaultMailLanguage].title && cleanedContents[defaultMailLanguage].body ? cleanedContents[defaultMailLanguage] : Object.values(cleanedContents).find((content) => content.title && content.body) ?? { title: "", body: "" };
      const response = await fetch("/local-api/mail-templates", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ id: template?.id, name: cleanName, title: primary.title, body: primary.body, contents: cleanedContents }) });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      onSaved();
    } catch (saveError) {
      setError(saveError instanceof Error ? `保存失败：${saveError.message}` : "保存失败");
    } finally {
      setSaving(false);
    }
  };
  return (
    <section className="mail-edit-page">
      <div className="mail-edit-card">
        <header>邮件模板</header>
        <div className="mail-template-name"><label>模板名称<input maxLength={MAX_TEMPLATE_NAME_LENGTH} value={name} onChange={(event) => setName(event.target.value)} placeholder="请输入模板名称" /></label><small>{name.length}/{MAX_TEMPLATE_NAME_LENGTH}</small></div>
        <div className="mail-language-tabs">{mailLanguages.map((language) => <button className={activeLanguage === language ? "active" : ""} key={language} onClick={() => setActiveLanguage(language)} type="button">{language}</button>)}</div>
        <div className="mail-form mail-template-form">
          <label className="mail-form-row"><span>邮件标题</span><input value={activeContent.title} onChange={(event) => updateContent({ title: event.target.value })} placeholder={`请输入${activeLanguage}邮件标题`} /></label>
          <label className="mail-form-row mail-template-body"><span>邮件内容</span><textarea value={activeContent.body} onChange={(event) => updateContent({ body: event.target.value })} placeholder={`请输入${activeLanguage}邮件内容`} /></label>
          <div className="mail-form-row">
            <span>模板文件</span>
            <div className="mail-template-file-actions">
              <button onClick={downloadTemplate} type="button">下载模板</button>
              <label className="mail-upload-link">上传Excel<input accept=".xlsx,.xls,.csv" onChange={(event) => {
                const file = event.target.files?.[0];
                event.currentTarget.value = "";
                if (file) void uploadTemplate(file);
              }} type="file" /></label>
            </div>
          </div>
          {error && <div className="mail-form-error">{error}</div>}
          <div className="mail-form-actions"><button disabled={saving} onClick={() => void save()} type="button">{saving ? "保存中..." : "保存"}</button><button disabled={saving} onClick={onBack} type="button">取消</button></div>
        </div>
      </div>
    </section>
  );
}

function GiftSuitePage({ active, canUploadItemTable, postWithToken }: { active: GiftSectionKey; canUploadItemTable: boolean; postWithToken: (endpoint: string, body: unknown) => Promise<ApiPostResponse> }) {
  const [rows, setRows] = React.useState<Record<string, unknown>[]>([]);
  const [view, setView] = React.useState<"list" | "edit">("list");
  const [items, setItems] = React.useState<ItemOption[]>([]);
  const [filterStatus, setFilterStatus] = React.useState("全部");
  const [codeQuery, setCodeQuery] = React.useState("");
  const [typeFilter, setTypeFilter] = React.useState("全部");
  const [status, setStatus] = React.useState("");

  const refreshItems = React.useCallback(async () => {
    const response = await fetch("/local-api/items");
    const payload = (await response.json().catch(() => ({}))) as { items?: ItemOption[] };
    setItems(payload.items ?? []);
  }, []);

  const refreshGiftList = React.useCallback(async () => {
    const result = await postWithToken("/gmGiftLst", {});
    const data = getObject(result.payload)?.data ?? getObject(result.payload);
    const list = getArray(getObject(data)?.Lst).filter((row): row is Record<string, unknown> => Boolean(getObject(row)));
    setRows(list);
  }, [postWithToken]);

  React.useEffect(() => {
    void refreshItems();
  }, [refreshItems]);

  React.useEffect(() => {
    if (active === "giftCode") void refreshGiftList().catch(() => setRows([]));
    setView("list");
  }, [active, refreshGiftList]);

  const uploadItemTable = async (file: File) => {
    const formData = new FormData();
    formData.append("file", file);
    const response = await fetch("/local-api/items/upload", { method: "POST", body: formData });
    const payload = (await response.json().catch(() => ({}))) as { items?: ItemOption[]; error?: string };
    if (!response.ok) throw new Error(payload.error || "上传道具表失败");
    setItems(payload.items ?? []);
    setStatus(`道具表已导入，共 ${payload.items?.length ?? 0} 个道具`);
  };

  if (active === "giftRecall") {
    return <GiftRecallPage />;
  }

  if (active === "giftClaim") {
    return <GiftClaimPage />;
  }

  if (view === "edit") {
    return <GiftEditor canUploadItemTable={canUploadItemTable} items={items} onBack={() => setView("list")} onSubmit={async (body) => {
      const result = await postWithToken("/gmGiftAdd", body);
      setStatus(result.ok ? "礼包码已保存" : `保存失败：HTTP ${result.status}`);
      setView("list");
      await refreshGiftList();
    }} onUploadItemTable={uploadItemTable} />;
  }

  const filteredRows = rows.filter((row) => {
    const id = formatCell(row.Id);
    const type = Number(row.Type) === 1 ? "global" : "personal";
    if (codeQuery.trim() && !id.includes(codeQuery.trim())) return false;
    if (typeFilter !== "全部" && type !== typeFilter) return false;
    if (filterStatus !== "全部") return true;
    return true;
  });

  return (
    <section className="gift-page">
      <div className="gift-filter-bar">
        <label>筛选状态：<select value={filterStatus} onChange={(event) => setFilterStatus(event.target.value)}><option>全部</option><option>启用</option><option>停用</option></select></label>
        <label>礼包码：<input value={codeQuery} onChange={(event) => setCodeQuery(event.target.value)} /></label>
        <button onClick={() => void refreshGiftList()} type="button"><Search size={14} />Search</button>
        <div className="gift-radio-filter"><RadioPill checked={typeFilter === "全部"} label="全部" onChange={() => setTypeFilter("全部")} /><RadioPill checked={typeFilter === "global"} label="global" onChange={() => setTypeFilter("global")} /><RadioPill checked={typeFilter === "personal"} label="personal" onChange={() => setTypeFilter("personal")} /></div>
      </div>
      <section className="gift-table-card">
        <button className="mail-primary-button" onClick={() => setView("edit")} type="button">新建</button>
        <MailDataTable
          columns={["礼包码", "模板名称", "类型", "最多可领", "已下载", "已领取", "是否启用", "生效时间", "过期时间", "操作"]}
          rows={filteredRows.map((row) => {
            const id = formatCell(row.Id);
            return {
              礼包码: id,
              模板名称: formatCell(row.Desc || row.Group || "自定义礼包"),
              类型: Number(row.Type) === 1 ? "global" : "personal",
              最多可领: formatCell(row.Num),
              已下载: "0",
              已领取: "0",
              是否启用: "是",
              生效时间: "暂无数据",
              过期时间: formatTimestamp(row.Et),
              操作: <div className="mail-action-buttons"><button type="button">查看</button><button onClick={() => setView("edit")} type="button">编辑</button><button onClick={() => void postWithToken("/gmGiftDel", { Id: id }).then(refreshGiftList)} type="button">删除</button></div>,
            };
          })}
        />
      </section>
      {status && <div className="mail-status">{status}</div>}
    </section>
  );
}

function GiftEditor({ canUploadItemTable, items, onBack, onSubmit, onUploadItemTable }: { canUploadItemTable: boolean; items: ItemOption[]; onBack: () => void; onSubmit: (body: unknown) => Promise<void>; onUploadItemTable: (file: File) => Promise<void> }) {
  const now = new Date();
  const later = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);
  const [code, setCode] = React.useState("");
  const [type, setType] = React.useState<"global" | "personal">("global");
  const [templateName, setTemplateName] = React.useState("");
  const [maxCount, setMaxCount] = React.useState("0");
  const [enabled, setEnabled] = React.useState(false);
  const [startTime, setStartTime] = React.useState(toDatetimeLocal(now));
  const [endTime, setEndTime] = React.useState(toDatetimeLocal(later));
  const [rewards, setRewards] = React.useState<MailRewardItem[]>([{ itemId: "", count: "0" }]);

  const submit = async () => {
    const ids = code.split(/[\n,，\s]+/).map((item) => item.trim()).filter(Boolean);
    const itemList = rewards.flatMap((reward) => {
      const itemId = Number(reward.itemId);
      const count = Number(reward.count);
      return Number.isFinite(itemId) && Number.isFinite(count) && itemId > 0 && count > 0 ? [itemId, count] : [];
    });
    const startSeconds = parseDatetimeLocalSeconds(startTime);
    const endSeconds = parseDatetimeLocalSeconds(endTime);
    if (!startSeconds || !endSeconds || endSeconds <= startSeconds) return;
    await onSubmit({ Id: ids.length ? ids : [code], Type: type === "global" ? 1 : 0, Group: Date.now() % 100000, Num: Number(maxCount), Et: endSeconds, ItemLst: itemList, Desc: templateName, Enabled: enabled, St: startSeconds });
  };

  return (
    <section className="gift-edit-page">
      <div className="gift-edit-card">
        <label className="gift-form-row"><span>礼包码</span><input value={code} onChange={(event) => setCode(event.target.value)} /><small>（字母、数字）</small></label>
        <div className="gift-form-row gift-radio-row"><span>类型</span><RadioPill checked={type === "global"} label="global" onChange={() => setType("global")} /><RadioPill checked={type === "personal"} label="personal" onChange={() => setType("personal")} /></div>
        <label className="gift-form-row"><span>邮件模板</span><input value={templateName} onChange={(event) => setTemplateName(event.target.value)} placeholder="请选择或填写模板名称" /></label>
        <label className="gift-form-row"><span>总兑换数</span><input value={maxCount} onChange={(event) => setMaxCount(event.target.value)} /></label>
        <label className="gift-form-row gift-check-row"><span>是否启用</span><input checked={enabled} onChange={(event) => setEnabled(event.target.checked)} type="checkbox" /></label>
        <RewardRows canUploadItemTable={canUploadItemTable} items={items} onUploadItemTable={onUploadItemTable} rewards={rewards} setRewards={setRewards} />
        <label className="gift-form-row"><span>生效时间</span><input type="datetime-local" value={startTime} onChange={(event) => setStartTime(event.target.value)} /><em>北京时间 {formatBeijingTime(startTime)}</em></label>
        <label className="gift-form-row"><span>过期时间</span><input type="datetime-local" value={endTime} onChange={(event) => setEndTime(event.target.value)} /><em>北京时间 {formatBeijingTime(endTime)}</em></label>
        <div className="gift-form-actions"><button onClick={() => void submit()} type="button">保存</button><button onClick={onBack} type="button">取消</button></div>
      </div>
    </section>
  );
}

function GiftClaimPage() {
  const [personalCode, setPersonalCode] = React.useState("");
  const [userId, setUserId] = React.useState("");
  return <section className="gift-page"><div className="gift-claim-block"><label>个人礼包码：<input value={personalCode} onChange={(event) => setPersonalCode(event.target.value)} /></label><button type="button"><Search size={14} />Search</button><MailDataTable columns={["礼包码", "个人礼包码", "模板名称", "类型", "兑换者", "兑换时间", "邮件奖励", "操作"]} rows={[]} /></div><div className="gift-claim-block"><label>userId：<input value={userId} onChange={(event) => setUserId(event.target.value)} /></label><button type="button"><Search size={14} />Search</button><MailDataTable columns={["礼包码", "个人礼包码", "模板名称", "类型", "兑换时间", "邮件奖励", "操作"]} rows={[]} /></div></section>;
}

function GiftRecallPage() {
  const [email, setEmail] = React.useState("");
  const [language, setLanguage] = React.useState("");
  return <section className="gift-page"><div className="gift-recall-bar"><label>邮箱：<input value={email} onChange={(event) => setEmail(event.target.value)} placeholder="请输入邮箱" /></label><label>语言：<select value={language} onChange={(event) => setLanguage(event.target.value)}><option value="">请选择语言</option>{mailLanguages.map((item) => <option key={item}>{item}</option>)}</select></label><button type="button">发送</button></div></section>;
}

function RewardTemplateList({ onCreate, onDelete, onEdit, templates }: { onCreate: () => void; onDelete: (template: RewardTemplate) => Promise<void>; onEdit: (template: RewardTemplate) => void; templates: RewardTemplate[] }) {
  return <section className="mail-page"><section className="mail-table-card"><button className="mail-primary-button" onClick={onCreate} type="button">新建</button><MailDataTable columns={["title", "创建时间", "更新时间", "操作"]} rows={templates.map((template) => ({ title: template.title, 创建时间: formatTimestampValue(template.createdAt), 更新时间: formatTimestampValue(template.updatedAt), 操作: <div className="mail-action-buttons"><button onClick={() => onEdit(template)} type="button">编辑</button><button onClick={() => void onDelete(template)} type="button">删除</button></div> }))} /></section></section>;
}

function RewardTemplateEditor({ canUploadItemTable, items, onBack, onSaved, onUploadItemTable, template }: { canUploadItemTable: boolean; items: ItemOption[]; onBack: () => void; onSaved: () => void; onUploadItemTable: (file: File) => Promise<void>; template?: RewardTemplate }) {
  const [title, setTitle] = React.useState(template?.title ?? "");
  const [rewards, setRewards] = React.useState<MailRewardItem[]>(template?.items ?? [{ itemId: "", count: "0" }]);
  const [error, setError] = React.useState("");
  const [saving, setSaving] = React.useState(false);
  const save = async () => {
    if (saving) return;
    const cleanTitle = title.trim();
    if (!cleanTitle) {
      setError("请输入奖励模板标题");
      return;
    }
    const rewardValidation = validateRewardRows(rewards, items);
    if (!rewardValidation.ok) {
      setError(rewardValidation.message ?? "请填写有效奖励");
      return;
    }
    if (!rewardValidation.itemList.length) {
      setError("请至少添加一个有效奖励");
      return;
    }
    setError("");
    setSaving(true);
    try {
      const response = await fetch("/local-api/reward-templates", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ id: template?.id, title: cleanTitle, items: filledRewards(rewards) }) });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      onSaved();
    } catch (saveError) {
      setError(saveError instanceof Error ? `保存失败：${saveError.message}` : "保存失败");
    } finally {
      setSaving(false);
    }
  };
  return <section className="mail-edit-page"><div className="mail-edit-card reward-template-card"><header>奖励模板</header><div className="mail-form"><label className="mail-form-row"><span>标题</span><input value={title} onChange={(event) => setTitle(event.target.value)} /></label><RewardRows canUploadItemTable={canUploadItemTable} items={items} onUploadItemTable={onUploadItemTable} rewards={rewards} setRewards={setRewards} />{error && <div className="mail-form-error">{error}</div>}<div className="mail-form-actions"><button disabled={saving} onClick={() => void save()} type="button">{saving ? "保存中..." : "保存"}</button><button disabled={saving} onClick={onBack} type="button">取消</button></div></div></div></section>;
}

function toDatetimeLocal(date: Date) {
  const local = new Date(date.getTime() - date.getTimezoneOffset() * 60000);
  return local.toISOString().slice(0, 16);
}

function parseDatetimeLocalSeconds(value: string) {
  const normalized = value.trim().replace(/\//g, "-").replace(" ", "T");
  const parsed = normalized ? new Date(normalized) : null;
  const seconds = parsed ? Math.floor(parsed.getTime() / 1000) : 0;
  return Number.isFinite(seconds) && seconds > 0 ? seconds : 0;
}

function formatBeijingTime(value: string) {
  if (!value) return "";
  const seconds = parseDatetimeLocalSeconds(value);
  return seconds ? formatTimestamp(seconds) : "";
}

function formatTimeDistance(value: string) {
  const seconds = parseDatetimeLocalSeconds(value);
  if (!seconds) return "时间无效";
  const diff = seconds * 1000 - Date.now();
  if (diff <= 0) return "已过期";
  const days = Math.floor(diff / 86400000);
  const hours = Math.floor((diff % 86400000) / 3600000);
  const minutes = Math.floor((diff % 3600000) / 60000);
  if (days > 0) return `距离过期 ${days}天${hours}小时`;
  if (hours > 0) return `距离过期 ${hours}小时${minutes}分钟`;
  return `距离过期 ${Math.max(1, minutes)}分钟`;
}

function formatTimestampValue(value: unknown) {
  if (typeof value === "string" && /^\d{4}-\d{2}-\d{2}/.test(value)) {
    const normalized = value.includes("T") ? value : value.replace(" ", "T");
    const timestampText = /(?:Z|[+-]\d{2}:?\d{2})$/.test(normalized) ? normalized : `${normalized}Z`;
    const parsed = new Date(timestampText);
    if (Number.isFinite(parsed.getTime())) return formatTimestamp(Math.floor(parsed.getTime() / 1000));
    return value.replace("T", " ").slice(0, 16);
  }
  return formatTimestamp(value);
}

function formatTimestamp(value: unknown) {
  const number = Number(value);
  if (!Number.isFinite(number) || number <= 0) return "暂无数据";
  return new Date(number * 1000).toLocaleString("zh-CN", {
    timeZone: "Asia/Shanghai",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  }).replace(/\//g, "-");
}

function secondsToDatetimeLocal(value: unknown) {
  const number = Number(value);
  return Number.isFinite(number) && number > 0 ? toDatetimeLocal(new Date(number * 1000)) : "";
}

function ActionCard({ action, isLoading, onSubmit }: { action: ApiAction; isLoading: boolean; onSubmit: (values: Record<string, string>) => void }) {
  const [values, setValues] = React.useState<Record<string, string>>({});
  return (
    <form className="action-card" onSubmit={(event) => { event.preventDefault(); onSubmit(values); }}>
      <div className="action-card-title"><strong>{action.label}</strong>{action.endpoint && <code>{action.endpoint}</code>}</div>
      {(action.fields ?? []).map((item) => (
        <label key={item.key}>{item.label}{item.kind === "textarea" ? <textarea value={values[item.key] ?? item.defaultValue ?? ""} onChange={(event) => setValues({ ...values, [item.key]: event.target.value })} placeholder={item.placeholder} /> : <input type={item.kind === "number" ? "number" : "text"} value={values[item.key] ?? item.defaultValue ?? ""} onChange={(event) => setValues({ ...values, [item.key]: event.target.value })} placeholder={item.placeholder} />}</label>
      ))}
      <button className="primary-button" disabled={isLoading} type="submit">{isLoading ? "请求中" : "提交"}</button>
    </form>
  );
}

function ResultFeed({ results }: { results: ApiResult[] }) {
  return (
    <section className="result-feed">
      <p>最近请求</p>
      {results.length === 0 ? <div className="empty-table"><strong>暂无请求记录</strong><span>提交接口后会在这里显示返回数据。</span></div> : results.map((item) => (
        <article className={`result-item ${item.ok ? "" : "fail"}`} key={item.id}>
          <header><strong>{item.label}</strong><span>{item.endpoint} / HTTP {item.status}</span></header>
          {apiPayloadSummary(item.payload) && <div className="result-summary">{apiPayloadSummary(item.payload)}</div>}
          <pre>{JSON.stringify(item.payload, null, 2)}</pre>
        </article>
      ))}
    </section>
  );
}

function UnavailablePanel({ module }: { module: ModuleConfig }) {
  return <section className="unavailable-panel"><module.icon size={34} /><strong>{module.title}暂未开放</strong><p>{module.description}</p></section>;
}

function NoticePage({ postWithToken }: { postWithToken: (endpoint: string, body: unknown) => Promise<ApiPostResponse> }) {
  const [notices, setNotices] = React.useState<NoticeConfig[]>([]);
  const [editing, setEditing] = React.useState<NoticeConfig | null>(null);
  const [form, setForm] = React.useState<NoticeConfig>({ slot: 1, title: "", body: "", imagePath: "", regBegin: "", regEnd: "", platforms: "", versions: "" });
  const [status, setStatus] = React.useState("");
  const palettes = ["blue", "green", "orange"];

  const refresh = React.useCallback(async () => {
    const result = await postWithToken("/gmNoticeLst", {});
    const error = apiBusinessError(result);
    if (error) {
      setStatus(`公告读取失败：${error}`);
      setNotices([1, 2, 3].map((slot) => ({ slot, title: "", body: "", imagePath: "" })));
      return;
    }
    setNotices(noticePayloadToConfigs(getApiData(result.payload)));
  }, [postWithToken]);

  React.useEffect(() => {
    void refresh();
  }, [refresh]);

  const openEditor = (notice?: NoticeConfig) => {
    const next = notice ?? notices[0] ?? { slot: 1, title: "", body: "", imagePath: "" };
    setEditing(next);
    setForm({ slot: next.slot, title: next.title ?? "", body: next.body ?? "", imagePath: next.imagePath ?? "", regBegin: next.regBegin ?? "", regEnd: next.regEnd ?? "", platforms: next.platforms ?? "", versions: next.versions ?? "" });
  };

  const save = async () => {
    const merged = [1, 2, 3].map((slot) => {
      const current = notices.find((notice) => Number(notice.slot) === slot) ?? { slot, title: "", body: "", imagePath: "" };
      return slot === form.slot ? form : current;
    });
    const result = await postWithToken("/gmNoticeAdd", configsToNoticePayload(merged));
    const error = apiBusinessError(result);
    if (error) {
      setStatus(`公告保存失败：${error}`);
      return;
    }
    setStatus("公告已保存到服务器");
    setEditing(null);
    await refresh();
  };

  return (
    <section className="notice-page">
      <h2>公告</h2>
      <button className="notice-edit-button" onClick={() => openEditor()} type="button"><Bell size={15} />添加/修改公告</button>
      <div className="notice-grid">
        {[1, 2, 3].map((slot, index) => {
          const notice = notices.find((item) => Number(item.slot) === slot) ?? { slot, title: "", body: "", imagePath: "" };
          return (
            <button className={`notice-card ${palettes[index]}`} key={slot} onClick={() => openEditor(notice)} type="button">
              <span className="notice-watermark">admin</span>
              <strong>公告 {slot}</strong>
              <h3>{notice.title || "未配置公告标题"}</h3>
              <div className="notice-body-box">{notice.body || "暂无公告内容"}</div>
              <label>配图路径</label>
              <div className="notice-image-path">{notice.imagePath || "未配置配图路径"}</div>
              <div className="tag-row">{notice.platforms && <small>系统：{notice.platforms}</small>}{notice.versions && <small>版本：{notice.versions}</small>}</div>
            </button>
          );
        })}
      </div>
      {status && <div className="mail-status">{status}</div>}
      {editing && (
        <div className="modal-backdrop" role="presentation">
          <section className="notice-modal" role="dialog" aria-modal="true">
            <header><strong>添加/修改公告</strong><button onClick={() => setEditing(null)} type="button">x</button></header>
            <div className="notice-form">
              <label>公告位置<select value={form.slot} onChange={(event) => setForm({ ...form, slot: Number(event.target.value) })}><option value={1}>公告 1</option><option value={2}>公告 2</option><option value={3}>公告 3</option></select></label>
              <label>公告标题<input value={form.title} onChange={(event) => setForm({ ...form, title: event.target.value })} placeholder="请输入公告标题" /></label>
              <label>公告内容<textarea value={form.body} onChange={(event) => setForm({ ...form, body: event.target.value })} placeholder="请输入公告内容" /></label>
              <label>配图路径<input value={form.imagePath} onChange={(event) => setForm({ ...form, imagePath: event.target.value })} placeholder="例如：/notice/banner_1.png" /></label>
              <label>注册开始<input type="datetime-local" value={form.regBegin ?? ""} onChange={(event) => setForm({ ...form, regBegin: event.target.value })} /></label>
              <label>注册结束<input type="datetime-local" value={form.regEnd ?? ""} onChange={(event) => setForm({ ...form, regEnd: event.target.value })} /></label>
              <label>系统筛选<input value={form.platforms ?? ""} onChange={(event) => setForm({ ...form, platforms: event.target.value })} placeholder="1=GooglePlay，2=iOS；留空全部" /></label>
              <label>版本筛选<input value={form.versions ?? ""} onChange={(event) => setForm({ ...form, versions: event.target.value })} placeholder="例如 1.8.0.0；留空全部" /></label>
              <footer><button onClick={() => void save()} type="button">保存</button><button onClick={() => setEditing(null)} type="button">取消</button></footer>
            </div>
          </section>
        </div>
      )}
    </section>
  );
}

function noticePayloadToConfigs(data: Record<string, unknown> | null): NoticeConfig[] {
  return [1, 2, 3].map((slot) => {
    const suffix = slot === 1 ? "" : String(slot);
    const platform = data?.[`Platform${slot}`];
    const version = data?.[`Version${slot}`];
    return {
      slot,
      title: String(data?.[`Titel${suffix}`] ?? ""),
      body: String(data?.[`Body${suffix}`] ?? ""),
      imagePath: String(data?.[`Rs${slot}`] ?? ""),
      regBegin: secondsToDatetimeLocal(data?.[`RegtBegin${slot}`]),
      regEnd: secondsToDatetimeLocal(data?.[`RegtEnd${slot}`]),
      platforms: Array.isArray(platform) ? platform.join(",") : "",
      versions: Array.isArray(version) ? version.join(",") : "",
    };
  });
}

function configsToNoticePayload(configs: NoticeConfig[]) {
  const payload: Record<string, unknown> = {};
  for (const config of configs) {
    const slot = Number(config.slot);
    const suffix = slot === 1 ? "" : String(slot);
    payload[`Titel${suffix}`] = config.title ?? "";
    payload[`Body${suffix}`] = config.body ?? "";
    payload[`Rs${slot}`] = config.imagePath ?? "";
    payload[`RegtBegin${slot}`] = config.regBegin ? parseDatetimeLocalSeconds(config.regBegin) : 0;
    payload[`RegtEnd${slot}`] = config.regEnd ? parseDatetimeLocalSeconds(config.regEnd) : 0;
    const platforms = toFlexibleNumberArray(config.platforms);
    const versions = toVersionNumberArray(config.versions);
    payload[`Platform${slot}`] = platforms.length ? platforms : null;
    payload[`Version${slot}`] = versions.length ? versions : null;
  }
  return payload;
}

function AccountPanel({ accounts, canManageAdmins, games, session, onAdd, onDelete, onUpdate, onClose }: { accounts: ManagedAccount[]; canManageAdmins: boolean; games: GameConfig[]; session: Session; onAdd: (account: ManagedAccount) => Promise<void>; onDelete: (accountId: number) => Promise<void>; onUpdate: (account: ManagedAccount) => Promise<void>; onClose: () => void }) {
  const currentGameKey = `${session.game}/${session.serverName}`;
  const emptyForm = { account: "", password: "", displayName: "", role: "运营", gameKeys: [currentGameKey], permissions: ["用户查询", "日志审计"], isManager: false };
  const [form, setForm] = React.useState(emptyForm);
  const [editingAccount, setEditingAccount] = React.useState<ManagedAccount | null>(null);
  const [formError, setFormError] = React.useState("");
  const gameOptions = React.useMemo(() => games.map((game) => ({ label: `${game.name} / ${game.serverName}`, value: `${game.name}/${game.serverName}` })), [games]);
  const beginEdit = (account: ManagedAccount) => {
    setEditingAccount(account);
    setForm({
      account: account.account,
      password: "",
      displayName: account.displayName,
      role: account.role,
      gameKeys: account.isManager ? gameOptions.map((game) => game.value) : account.games,
      permissions: account.permissions,
      isManager: Boolean(account.isManager),
    });
    setFormError("");
  };
  const resetForm = () => {
    setEditingAccount(null);
    setForm(emptyForm);
    setFormError("");
  };

  React.useEffect(() => {
    const available = new Set(gameOptions.map((game) => game.value));
    const filtered = form.gameKeys.filter((game) => available.has(game));
    if (filtered.length !== form.gameKeys.length || (!filtered.length && gameOptions[0])) {
      setForm((current) => ({ ...current, gameKeys: filtered.length ? filtered : gameOptions[0] ? [gameOptions[0].value] : [] }));
    }
  }, [form.gameKeys, gameOptions]);

  return (
    <div className="modal-backdrop" role="presentation">
      <section className="account-panel" role="dialog" aria-modal="true">
        <header><div><strong>账号管理</strong><span>创建成员账号、分配游戏和功能权限</span></div><button onClick={onClose} type="button">x</button></header>
        <div className="account-layout">
          <form className="account-form" onSubmit={async (event) => {
            event.preventDefault();
            try {
              setFormError("");
              const payload = { id: editingAccount?.id ?? Date.now(), account: form.account, password: form.password, displayName: form.displayName || form.account, role: form.role, games: form.isManager ? gameOptions.map((game) => game.value) : form.gameKeys, permissions: form.permissions, isManager: canManageAdmins && form.isManager, status: editingAccount?.status ?? "启用" };
              if (editingAccount) {
                await onUpdate(payload);
                resetForm();
              } else {
                await onAdd(payload);
                setForm((current) => ({ ...current, account: "", password: "", displayName: "", gameKeys: [currentGameKey], isManager: false }));
              }
            } catch (error) {
              setFormError(error instanceof Error ? error.message : editingAccount ? "保存账号失败" : "创建账号失败");
            }
          }}>
            <label>登录账号<input disabled={Boolean(editingAccount)} value={form.account} onChange={(event) => setForm({ ...form, account: event.target.value })} placeholder="成员账号" /></label>
            <label>登录密码<input type="password" value={form.password} onChange={(event) => setForm({ ...form, password: event.target.value })} placeholder={editingAccount ? "不修改密码可留空" : "成员密码"} /></label>
            <label>显示名称<input value={form.displayName} onChange={(event) => setForm({ ...form, displayName: event.target.value })} placeholder="成员姓名" /></label>
            <label>角色<select value={form.role} onChange={(event) => setForm({ ...form, role: event.target.value })}><option>策划</option><option>程序</option><option>测试</option><option>运营</option></select></label>
            <label>游戏权限<MultiSelectDropdown options={gameOptions} values={form.gameKeys} onChange={(gameKeys) => setForm({ ...form, gameKeys })} placeholder="请选择游戏区服" /></label>
            {canManageAdmins && <label className="permission-check admin-toggle"><input checked={form.isManager} onChange={(event) => setForm({ ...form, isManager: event.target.checked })} type="checkbox" />设为后台管理员</label>}
            <label>功能权限<MultiSelectDropdown options={permissionOptions.map((item) => ({ label: item, value: item }))} values={form.permissions} onChange={(permissions) => setForm({ ...form, permissions })} placeholder="请选择功能权限" /></label>
            <div className="form-button-row">
              <button className="primary-button" type="submit"><Plus size={15} />{editingAccount ? "保存修改" : "创建账号"}</button>
              {editingAccount && <button onClick={resetForm} type="button">取消编辑</button>}
            </div>
            {formError && <div className="form-error">{formError}</div>}
          </form>
          <div className="managed-list">{accounts.length === 0 ? <div className="empty-table"><strong>暂无子账号</strong><span>请在左侧创建账号给其他人使用。</span></div> : accounts.map((account) => (
            <article className="managed-account" key={account.id}>
              <div><strong>{account.displayName}</strong><span>{account.account} / {account.role}{account.isManager ? " / 后台管理员" : ""}</span></div>
              <div className="tag-row">{account.isManager ? <small>全部游戏</small> : account.games.map((item) => <small key={item}>{item}</small>)}{account.permissions.slice(0, 3).map((item) => <small key={item}>{item}</small>)}</div>
              <div className="managed-actions"><button onClick={() => beginEdit(account)} type="button"><UserCog size={14} />编辑</button>{canManageAdmins ? <button onClick={() => void onDelete(account.id)} type="button"><Trash2 size={14} />删除</button> : <button disabled type="button"><Trash2 size={14} />仅admin可删</button>}</div>
            </article>
          ))}</div>
        </div>
      </section>
    </div>
  );
}

function MultiSelectDropdown({ onChange, options, placeholder, values }: { onChange: (values: string[]) => void; options: Array<{ label: string; value: string }>; placeholder: string; values: string[] }) {
  const [open, setOpen] = React.useState(false);
  const selected = options.filter((option) => values.includes(option.value));
  const toggle = (value: string) => {
    onChange(values.includes(value) ? values.filter((item) => item !== value) : [...values, value]);
  };

  return (
    <div className="multi-select">
      <button className="multi-select-trigger" onClick={() => setOpen((current) => !current)} type="button">
        <span>{selected.length ? selected.map((option) => option.label).join("、") : placeholder}</span>
        <ChevronDown size={15} />
      </button>
      {open && (
        <div className="multi-select-menu">
          {options.map((option) => (
            <label className="multi-select-option" key={option.value}>
              <input checked={values.includes(option.value)} onChange={() => toggle(option.value)} type="checkbox" />
              {option.label}
            </label>
          ))}
        </div>
      )}
    </div>
  );
}

function UserLogPanel({ onClose }: { onClose: () => void }) {
  const [logs, setLogs] = React.useState<UserLog[]>([]);
  const [query, setQuery] = React.useState("");
  const [activeCategory, setActiveCategory] = React.useState("全部");
  const categories = ["全部", "登录", "礼包码", "邮件", "封号", "禁言", "账号", "区服", "玩家", "系统"];

  const refresh = React.useCallback(async () => {
    const response = await fetch("/local-api/user-logs");
    const payload = (await response.json().catch(() => ({}))) as { logs?: UserLog[] };
    setLogs(payload.logs ?? []);
  }, []);

  React.useEffect(() => {
    void refresh();
  }, [refresh]);

  const categoryOf = (log: UserLog) => {
    const text = `${log.action} ${log.target} ${log.detail ?? ""}`;
    if (text.includes("登录") || text.includes("进入区服")) return "登录";
    if (text.includes("gmGift") || text.includes("礼包")) return "礼包码";
    if (text.includes("gmMail") || text.includes("邮件")) return "邮件";
    if (text.includes("gmStateAdd") || text.includes("封号") || text.includes("解封")) return "封号";
    if (text.includes("禁言") || text.includes("解禁")) return "禁言";
    if (text.includes("账号")) return "账号";
    if (text.includes("区服") || text.includes("游戏区服")) return "区服";
    if (text.includes("gmPlayer") || text.includes("gmGetAcount") || text.includes("gmSetAcount") || text.includes("玩家") || text.includes("UID")) return "玩家";
    return "系统";
  };

  const categoryCounts = categories.reduce<Record<string, number>>((total, category) => {
    total[category] = category === "全部" ? logs.length : logs.filter((log) => categoryOf(log) === category).length;
    return total;
  }, {});

  const filtered = logs.filter((log) => {
    if (activeCategory !== "全部" && categoryOf(log) !== activeCategory) return false;
    const keyword = query.trim();
    if (!keyword) return true;
    return [log.operator, log.action, log.target, log.game, log.serverName, log.detail, log.result].some((value) => String(value ?? "").includes(keyword));
  });

  return (
    <div className="modal-backdrop" role="presentation">
      <section className="user-log-panel" role="dialog" aria-modal="true">
        <header><div><strong>用户日志</strong><span>查看所有用户登录、接口请求和管理操作记录</span></div><button onClick={onClose} type="button">x</button></header>
        <div className="user-log-toolbar">
          <label>搜索<input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="账号 / 动作 / 目标" /></label>
          <button onClick={() => void refresh()} type="button">刷新</button>
        </div>
        <div className="user-log-tabs">
          {categories.map((category) => <button className={activeCategory === category ? "active" : ""} key={category} onClick={() => setActiveCategory(category)} type="button">{category}<span>{categoryCounts[category] ?? 0}</span></button>)}
        </div>
        <div className="user-log-table-wrap">
          <table className="user-log-table">
            <thead><tr><th>时间</th><th>账号</th><th>角色</th><th>动作</th><th>目标</th><th>游戏区服</th><th>结果</th><th>详情</th></tr></thead>
            <tbody>
              {filtered.length ? filtered.map((log) => (
                <tr key={log.id}>
                  <td>{log.time}</td>
                  <td>{log.operator}</td>
                  <td>{log.role}</td>
                  <td>{log.action}</td>
                  <td>{log.target}</td>
                  <td>{log.game && log.serverName ? `${log.game}/${log.serverName}` : "暂无数据"}</td>
                  <td><span className={log.result === "成功" ? "log-success" : "log-fail"}>{log.result}</span></td>
                  <td>{log.detail || "暂无数据"}</td>
                </tr>
              )) : <tr><td colSpan={8}>暂无日志</td></tr>}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  );
}

function GamePanel({ games, onAdd, onDelete, onUpdate, onClose }: { games: GameConfig[]; onAdd: (game: GameConfig) => Promise<void>; onDelete: (gameId: number) => Promise<void>; onUpdate: (game: GameConfig) => Promise<void>; onClose: () => void }) {
  const emptyForm = { name: "包包4", serverName: "", serverUrl: "", environment: portal === "prod" ? "Prod" : "Test", logo: "包", serverAccount: "", serverPassword: "", iconUrl: "", backgroundUrl: "" };
  const [form, setForm] = React.useState<GameConfig>(emptyForm);
  const [editingId, setEditingId] = React.useState<number | null>(null);
  const [formError, setFormError] = React.useState("");
  const loadImage = (file: File | undefined, key: "iconUrl" | "backgroundUrl") => {
    if (!file) return;
    if (!file.type.startsWith("image/")) {
      setFormError("请上传图片文件");
      return;
    }
    const reader = new FileReader();
    reader.onload = () => setForm((current) => ({ ...current, [key]: String(reader.result ?? "") }));
    reader.readAsDataURL(file);
  };
  const resetForm = () => {
    setEditingId(null);
    setForm(emptyForm);
    setFormError("");
  };
  const beginEdit = (game: GameConfig) => {
    setEditingId(game.id ?? null);
    setForm({ ...emptyForm, ...game, iconUrl: game.iconUrl ?? "", backgroundUrl: game.backgroundUrl ?? "" });
    setFormError("");
  };

  return (
    <div className="modal-backdrop" role="presentation">
      <section className="account-panel" role="dialog" aria-modal="true">
        <header><div><strong>游戏 / 区服管理</strong><span>{editingId ? "修改已有游戏区服配置" : "添加或删除当前入口可访问的游戏区服"}</span></div><button onClick={onClose} type="button">x</button></header>
        <div className="account-layout">
          <form className="account-form" onSubmit={async (event) => {
            event.preventDefault();
            try {
              setFormError("");
              if (!form.serverAccount?.trim() || !form.serverPassword?.trim()) {
                setFormError("请填写该区服的服务端账号和密码");
                return;
              }
              if (editingId) {
                await onUpdate({ ...form, id: editingId });
                resetForm();
              } else {
                await onAdd(form);
                setForm((current) => ({ ...emptyForm, name: current.name }));
              }
            } catch (error) {
              setFormError(error instanceof Error ? error.message : editingId ? "保存区服失败" : "新增区服失败");
            }
          }}>
            <label>游戏名<input value={form.name} onChange={(event) => setForm({ ...form, name: event.target.value })} placeholder="例如：包包4" /></label>
            <label>区服<input value={form.serverName} onChange={(event) => setForm({ ...form, serverName: event.target.value })} placeholder="例如：测试服、开发服、正式服" /></label>
            <label>服务端地址<input value={form.serverUrl} onChange={(event) => setForm({ ...form, serverUrl: event.target.value })} placeholder="http://52.77.195.98:9089" /></label>
            <label>服务端账号<input autoComplete="off" value={form.serverAccount ?? ""} onChange={(event) => setForm({ ...form, serverAccount: event.target.value })} placeholder="请输入该区服服务端账号" /></label>
            <label>服务端密码<input autoComplete="new-password" type="password" value={form.serverPassword ?? ""} onChange={(event) => setForm({ ...form, serverPassword: event.target.value })} placeholder="请输入该区服服务端密码" /></label>
            <label>游戏Icon图片<input accept="image/*" type="file" onChange={(event) => loadImage(event.target.files?.[0], "iconUrl")} /></label>
            <label>游戏背景图片<input accept="image/*" type="file" onChange={(event) => loadImage(event.target.files?.[0], "backgroundUrl")} /></label>
            {(form.iconUrl || form.backgroundUrl) && (
              <div className="game-asset-preview" style={form.backgroundUrl ? { backgroundImage: `linear-gradient(180deg, rgb(255 255 255 / 60%), rgb(255 255 255 / 92%)), url(${form.backgroundUrl})` } : undefined}>
                <div className="game-logo">{form.iconUrl ? <img alt="游戏Icon预览" src={form.iconUrl} /> : form.logo}</div>
                <span>{form.name || "游戏名"} / {form.serverName || "区服"}</span>
              </div>
            )}
            <div className="form-button-row">
              <button className="primary-button" type="submit"><Plus size={15} />{editingId ? "保存修改" : "新增区服"}</button>
              {editingId && <button onClick={resetForm} type="button">取消编辑</button>}
            </div>
            {formError && <div className="form-error">{formError}</div>}
          </form>
          <div className="managed-list">{games.length === 0 ? <div className="empty-table"><strong>暂无游戏区服</strong><span>请在左侧新增。</span></div> : games.map((game) => (
            <article className="managed-account" key={game.id ?? `${game.name}/${game.serverName}`}>
              <div><strong>{game.name} / {game.serverName}</strong><span>{game.serverAccount ? "已配置服务端账号" : "使用默认服务端账号"}</span></div>
              <div className="tag-row"><small>{game.serverUrl}</small><small>{game.iconUrl ? "已上传Icon" : "未上传Icon"}</small>{game.backgroundUrl && <small>已上传背景</small>}</div>
              <div className="managed-actions"><button onClick={() => beginEdit(game)} type="button"><UserCog size={14} />编辑</button><button onClick={() => game.id && void onDelete(game.id)} type="button"><Trash2 size={14} />删除</button></div>
            </article>
          ))}</div>
        </div>
      </section>
    </div>
  );
}
ReactDOM.createRoot(document.getElementById("root")!).render(<App />);


