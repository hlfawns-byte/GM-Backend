import React from "react";
import ReactDOM from "react-dom/client";
import * as XLSX from "xlsx";
import {
  Bell,
  ChevronDown,
  CreditCard,
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
  | "noticeTemplate"
  | "sprint"
  | "systemMsg"
  | "serverTime"
  | "orderRefund"
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
  templateName?: string;
  title: string;
  body: string;
  contents?: Record<string, MailTemplateContent>;
  imagePath: string;
  typ?: number;
  sid?: string;
  regBegin?: string;
  regEnd?: string;
  platforms?: string;
  versions?: string;
  conditions?: ConditionRow[];
  updatedAt?: string;
};

type ConditionRow = {
  id: number;
  field: string;
  op: string;
  value: string;
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
const NOTICE_DEFAULT_REG_BEGIN = "2020-01-01T00:00";
const NOTICE_DEFAULT_REG_END = "2050-12-31T23:59";
const NOTICE_DEFAULT_VERSION_RANGE = "0,4294967295";
const NOTICE_DEFAULT_IMAGE = "notice_bg_1";

function defaultRegEndTime() {
  return toDatetimeLocal(new Date());
}

const languageDefinitions = [
  { id: 1, label: "中文(简体)", code: "ChineseSimplified", aliases: ["中文(简体)", "简体中文", "ChineseSimplified"] },
  { id: 2, label: "中文(繁体)", code: "ChineseTraditional", aliases: ["中文(繁体)", "繁体中文", "ChineseTraditional"] },
  { id: 3, label: "英文", code: "English", aliases: ["英文", "英语", "English"] },
  { id: 4, label: "日语", code: "Japanese", aliases: ["日文", "日语", "Japanese"] },
  { id: 5, label: "韩语", code: "Korean", aliases: ["韩文", "韩语", "Korean"] },
  { id: 6, label: "俄语", code: "Russian", aliases: ["俄语", "Russian"] },
  { id: 7, label: "越南语", code: "Vietnamese", aliases: ["越南语", "Vietnamese"] },
  { id: 8, label: "德语", code: "German", aliases: ["德语", "German"] },
  { id: 9, label: "葡萄牙语", code: "Portuguese", aliases: ["葡萄牙语", "Portuguese", "葡语"] },
  { id: 10, label: "西班牙语", code: "Spanish", aliases: ["西班牙语", "Spanish"] },
  { id: 11, label: "法语", code: "French", aliases: ["法语", "French"] },
];
const mailLanguages = languageDefinitions.map((language) => language.label);
const defaultLanguageDefinition = languageDefinitions.find((language) => language.id === 3) ?? languageDefinitions[0];
const defaultMailLanguage = defaultLanguageDefinition.label;

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
      { key: "noticeTemplate", label: "公告模板" },
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
      { key: "orderRefund", label: "三方支付退款" },
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
  noticeTemplate: { title: "公告模板", description: "维护公告标题和内容模板", icon: Bell, status: "live", actions: [] },
  sprint: { title: "冲刺活动", description: "冲刺活动新增和列表", icon: Trophy, status: "live", actions: [{ key: "gmSprintLst", label: "活动列表", endpoint: "/gmSprintLst", fields: [], buildBody: () => ({}) }, { key: "gmSprintAddLst", label: "新增活动", endpoint: "/gmSprintAddLst", fields: [{ key: "Body", label: "活动JSON", kind: "textarea" }], buildBody: (v) => parseJson(v.Body) }] },
  systemMsg: { title: "系统消息", description: "向在线玩家发送系统消息", icon: MessageSquare, status: "live", actions: [{ key: "gmSendSystemMsg", label: "发送消息", endpoint: "/gmSendSystemMsg", fields: [{ key: "Msg", label: "消息内容", kind: "textarea" }], buildBody: (v) => ({ Msg: v.Msg }) }] },
  serverTime: { title: "开服管理", description: "查询和配置开服时间", icon: Server, status: "live", actions: [] },
  orderRefund: { title: "三方支付退款", description: "输入三方支付订单号处理用户退款", icon: CreditCard, status: "live", actions: [] },
  chatClear: { title: "清理聊天", description: "清理聊天消息", icon: Send, status: "live", actions: [{ key: "gmClearChatMsg", label: "清理聊天", endpoint: "/gmClearChatMsg", fields: [], buildBody: () => ({}) }] },
  rank: { title: "排行榜查询", description: "当前接口文档暂未开放", icon: Trophy, status: "pending", actions: [] },
  logs: { title: "作业日志", description: "当前接口文档暂未开放", icon: Database, status: "pending", actions: [] },
};

function toNumberArray(value?: string) {
  return String(value ?? "")
    .split(/[\s,，;；]+/)
    .map((item) => item.trim())
    .filter(Boolean)
    .map((item) => Number(item.trim()))
    .filter((item) => Number.isFinite(item));
}

function toFlexibleNumberArray(value?: string) {
  return String(value ?? "")
    .split(/[\s,，;；]+/)
    .map((item) => item.trim())
    .filter(Boolean)
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

function versionTextToCode(value: string) {
  const parts = parseVersionParts(value);
  if (!parts) return null;
  return versionPartsToCode(parts);
}

function parseVersionParts(value: string) {
  const text = value.trim();
  if (!/^\d+(?:\.\d+){1,3}$/.test(text)) return null;
  const parts = text.split(".").map((part) => Number(part));
  if (!parts.every((part, index) => Number.isInteger(part) && part >= 0 && part <= (index >= 2 ? 20 : 99))) return null;
  return parts;
}

function versionInputToParts(value: string, op: string) {
  const parts = parseVersionParts(value);
  if (!parts) return null;
  if (parts.length === 2 && op === "<=") return [parts[0], parts[1], 0, 0];
  if (parts.length === 2) return [parts[0], parts[1], 0];
  return parts;
}

function versionPartsToCode(parts: number[]) {
  return parts.reduce((total, part, index) => total + part * Math.pow(100, parts.length - index - 1), 0);
}

function compareVersionParts(left: number[], right: number[]) {
  const maxLength = Math.max(left.length, right.length, 3);
  for (let index = 0; index < maxLength; index += 1) {
    const diff = (left[index] ?? 0) - (right[index] ?? 0);
    if (diff !== 0) return diff;
  }
  return 0;
}

function versionLowerBoundParts(target: number[]) {
  const major = target[0] ?? 1;
  const minor = Math.max(0, (target[1] ?? 0) - 2);
  return [major, minor, 0, 0];
}

function enumerateVersionsUpTo(target: number[], includeTarget: boolean) {
  const maxGeneratedPart = 20;
  const result: number[] = [];
  const targetMajor = target[0] ?? 1;
  const lowerBound = versionLowerBoundParts(target);
  for (let major = lowerBound[0]; major <= targetMajor; major += 1) {
    const minMinor = major === lowerBound[0] ? lowerBound[1] : 0;
    const maxMinor = major === targetMajor ? target[1] ?? 0 : 99;
    for (let minor = minMinor; minor <= maxMinor; minor += 1) {
      const maxPatch = major === targetMajor && minor === maxMinor ? target[2] ?? 0 : maxGeneratedPart;
      for (let patch = 0; patch <= maxPatch; patch += 1) {
        const base = [major, minor, patch];
        const baseCompare = compareVersionParts(base, target);
        if (baseCompare < 0 || (includeTarget && baseCompare === 0)) {
          result.push(versionPartsToCode(base));
          result.push(versionPartsToCode([major, minor, patch, 0]));
        }
        const maxHotfix = major === targetMajor && minor === maxMinor && patch === maxPatch ? target[3] ?? 0 : maxGeneratedPart;
        for (let hotfix = 1; hotfix <= maxHotfix; hotfix += 1) {
          const hotfixParts = [major, minor, patch, hotfix];
          const hotfixCompare = compareVersionParts(hotfixParts, target);
          if (hotfixCompare < 0 || (includeTarget && hotfixCompare === 0)) result.push(versionPartsToCode(hotfixParts));
        }
      }
    }
  }
  return result;
}

function toVersionConditionArray(rows: Array<{ field: string; op: string; value: string }>) {
  const versionRows = rows.filter((row) => row.field === "version" && row.value.trim());
  if (!versionRows.length) return [];
  let selected: number[] | null = null;
  for (const row of versionRows) {
    const partsList = String(row.value ?? "")
      .split(/[\s,，;；]+/)
      .map((item) => versionInputToParts(item, row.op))
      .filter((item): item is number[] => Boolean(item));
    if (!partsList.length) return [];
    let rowCodes: number[] = [];
    if (row.op === "=") {
      rowCodes = partsList.map(versionPartsToCode);
    } else if (row.op === "<=" || row.op === "<") {
      rowCodes = Array.from(new Set(partsList.flatMap((parts) => enumerateVersionsUpTo(parts, row.op === "<="))));
    } else {
      return [];
    }
    const rowSet = new Set<number>(rowCodes);
    selected = selected ? selected.filter((code) => rowSet.has(code)) : rowCodes;
  }
  return Array.from(new Set(selected ?? [])).sort(compareVersionCodes);
}

function versionCodeToText(value: unknown) {
  let code = Number(value);
  if (!Number.isFinite(code) || code < 0) return String(value ?? "");
  code = Math.floor(code);
  if (code >= 100_000_000 && code % 100 === 1) code = Math.floor(code / 100);
  const parts = [0, 0, 0, 0];
  for (let index = 3; index >= 0; index -= 1) {
    parts[index] = code % 100;
    code = Math.floor(code / 100);
  }
  while (parts.length > 1 && parts[0] === 0) parts.shift();
  return parts.join(".");
}

function versionCodeToParts(value: unknown) {
  return versionCodeToText(value)
    .split(".")
    .map((part) => Number(part))
    .filter((part) => Number.isFinite(part));
}

function compareVersionCodes(left: unknown, right: unknown) {
  return compareVersionParts(versionCodeToParts(left), versionCodeToParts(right));
}

function formatVersionConditionList(values: unknown[]) {
  if (values.length > 30) {
    const numbers = values.map((value) => Number(value)).filter((value) => Number.isFinite(value)).sort(compareVersionCodes);
    if (numbers.length) return [`${versionCodeToText(numbers[0])} - ${versionCodeToText(numbers[numbers.length - 1])}`];
  }
  const formatted: string[] = [];
  for (let index = 0; index < values.length; index += 1) {
    const current = Number(values[index]);
    const next = Number(values[index + 1]);
    if (Number.isFinite(current) && Number.isFinite(next) && next === current * 100 + 1) {
      formatted.push(versionCodeToText(current));
      index += 1;
    } else {
      formatted.push(versionCodeToText(values[index]));
    }
  }
  return formatted;
}

function formatVersionConditionRows(rows: unknown[]) {
  return rows
    .map((row) => getObject(row))
    .filter((row): row is Record<string, unknown> => Boolean(row))
    .filter((row) => String(row.field ?? "") === "version" && String(row.value ?? "").trim())
    .flatMap((row) => {
      const op = String(row.op ?? "=");
      return String(row.value ?? "")
        .split(/[\s,，;；]+/)
        .map((value) => value.trim())
        .filter(Boolean)
        .map((value) => `${op}${value}`);
    });
}

function toNoticeVersionArray(value?: string) {
  const text = String(value ?? "").trim();
  if (!text) return [];
  return text.includes(".") ? toVersionNumberArray(text) : toFlexibleNumberArray(text);
}

function defaultMailRegEndDisplay() {
  return formatBeijingTime(defaultRegEndTime());
}

function toPlatformNumberArray(value?: string) {
  return String(value ?? "")
    .split(/[\s,，;；]+/)
    .map((item) => Number(item.trim()))
    .filter((item) => item === 1 || item === 2);
}

function gameServerDisplayName(id: unknown) {
  const number = Number(id);
  if (!Number.isFinite(number) || number <= 0) return String(id ?? "");
  return number <= 123 ? `GL-${number}` : `TK-${number - 123}`;
}

function dateToDatetimeLocal(value: string, endOfDay = false) {
  if (!value) return "";
  if (value.includes("T")) return value;
  return `${value}T${endOfDay ? "23:59" : "00:00"}`;
}

function templatePrimaryContent(template: MailTemplate) {
  const contents = normalizeLanguageContents(template.contents, { title: template.title, body: template.body });
  return contents[defaultMailLanguage] ?? Object.values(contents).find((content) => content.title?.trim() || content.body?.trim()) ?? { title: template.title, body: template.body };
}

function stringFromCell(value: unknown) {
  return String(value ?? "").trim();
}

function matchLanguage(value: unknown) {
  const text = stringFromCell(value).toLowerCase();
  if (!text) return undefined;
  return languageDefinitions.find((language) => String(language.id) === text || language.label.toLowerCase() === text || language.code.toLowerCase() === text || language.aliases.some((alias) => alias.toLowerCase() === text));
}

function emptyLanguageContents() {
  return Object.fromEntries(mailLanguages.map((language) => [language, { title: "", body: "" }])) as Record<string, MailTemplateContent>;
}

function normalizeLanguageContents(contents?: Record<string, MailTemplateContent>, fallback?: MailTemplateContent) {
  const normalized = emptyLanguageContents();
  for (const language of languageDefinitions) {
    const sourceKey = [language.label, language.code, ...language.aliases].find((key) => contents?.[key]);
    if (sourceKey && contents?.[sourceKey]) {
      normalized[language.label] = { title: contents[sourceKey].title ?? "", body: contents[sourceKey].body ?? "" };
    }
  }
  if (fallback && !normalized[defaultMailLanguage].title && !normalized[defaultMailLanguage].body) {
    normalized[defaultMailLanguage] = { title: fallback.title ?? "", body: fallback.body ?? "" };
  }
  return normalized;
}

function languageContentList(contents?: Record<string, MailTemplateContent>) {
  const normalized = normalizeLanguageContents(contents);
  return languageDefinitions.map((language) => ({
    id: language.id,
    language: language.label,
    code: language.code,
    title: normalized[language.label]?.title ?? "",
    body: normalized[language.label]?.body ?? "",
  }));
}

function fillMissingLanguageContents(contents?: Record<string, MailTemplateContent>, fallback?: MailTemplateContent) {
  const normalized = normalizeLanguageContents(contents, fallback);
  const defaultContent = normalized[defaultMailLanguage];
  const fallbackContent = defaultContent.title && defaultContent.body ? defaultContent : Object.values(normalized).find((content) => content.title && content.body) ?? { title: fallback?.title ?? "", body: fallback?.body ?? "" };
  return Object.fromEntries(languageDefinitions.map((language) => {
    const content = normalized[language.label] ?? { title: "", body: "" };
    return [language.label, {
      title: content.title.trim() || fallbackContent.title,
      body: content.body.trim() || fallbackContent.body,
    }];
  })) as Record<string, MailTemplateContent>;
}

function multiLanguagePayload(contents?: Record<string, MailTemplateContent>) {
  const languages = languageContentList(fillMissingLanguageContents(contents));
  return {
    type: "multiLanguage",
    defaultLanguageId: defaultLanguageDefinition.id,
    languages,
    contents: Object.fromEntries(languages.map((language) => [language.language, { title: language.title, body: language.body }])),
  };
}

function mailLangListPayload(contents?: Record<string, MailTemplateContent>) {
  return languageContentList(fillMissingLanguageContents(contents))
    .filter((language) => language.id !== defaultLanguageDefinition.id && language.title.trim() && language.body.trim())
    .map((language) => ({
      Language: language.id,
      Titel: language.title,
      Body: language.body,
    }));
}

async function parseMailTemplateFile(file: File) {
  const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: false });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) throw new Error("模板文件中没有可读取的工作表");
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "" });
  const usefulRows = rows.filter((row) => row.some((cell) => stringFromCell(cell)));
  if (!usefulRows.length) throw new Error("模板文件没有内容");
  const contents = Object.fromEntries(mailLanguages.map((language) => [language, { title: "", body: "" }])) as Record<string, MailTemplateContent>;
  const threeRowLanguages = usefulRows[0] ?? [];
  const threeRowTitles = usefulRows[1] ?? [];
  const threeRowBodies = usefulRows[2] ?? [];
  const threeRowLanguageMatches = threeRowLanguages
    .map((cell, index) => ({ language: matchLanguage(cell)?.label, index }))
    .filter((item): item is { language: string; index: number } => Boolean(item.language));
  if (threeRowLanguageMatches.length >= 2 && (stringFromCell(threeRowTitles[0]).toLowerCase() === "title" || stringFromCell(threeRowBodies[0]).toLowerCase() === "desc")) {
    for (const item of threeRowLanguageMatches) {
      contents[item.language] = { title: stringFromCell(threeRowTitles[item.index]), body: stringFromCell(threeRowBodies[item.index]) };
    }
    if (!Object.values(contents).some((content) => content.title || content.body)) throw new Error("模板文件没有邮件标题或邮件内容");
    const primary = contents[defaultMailLanguage].title || contents[defaultMailLanguage].body ? contents[defaultMailLanguage] : Object.values(contents).find((content) => content.title || content.body) ?? { title: "", body: "" };
    return { name: "", title: primary.title, body: primary.body, contents };
  }
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
  if (languageColumn >= 0) {
    for (const item of dataRows) {
      const language = matchLanguage(item[languageColumn])?.label;
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
  const normalized = value.trim().toUpperCase();
  const glMatch = normalized.match(/^GL-?(\d+)$/);
  if (glMatch) return Number(glMatch[1]);
  const tkMatch = normalized.match(/^TK-?(\d+)$/);
  if (tkMatch) return Number(tkMatch[1]) + 123;
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

function chunkArray<T>(items: T[], size: number) {
  const chunks: T[][] = [];
  for (let index = 0; index < items.length; index += size) chunks.push(items.slice(index, index + size));
  return chunks;
}

function mailGroupIds(row?: Record<string, unknown>) {
  if (!row) return [];
  const ids = [row.Id ?? row.id, ...getArray(row.__childMailIds)]
    .map((value) => String(value ?? "").trim())
    .filter(Boolean);
  return Array.from(new Set(ids));
}

function getApiData(payload: unknown) {
  return getObject(getObject(payload)?.data) ?? getObject(payload);
}

function parseTimeSeconds(value: unknown) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return Number.isFinite(value) ? Math.floor(value > 10_000_000_000 ? value / 1000 : value) : 0;
  const text = String(value).trim();
  if (!text) return 0;
  const numeric = Number(text);
  if (Number.isFinite(numeric)) return Math.floor(numeric > 10_000_000_000 ? numeric / 1000 : numeric);
  const parsed = new Date(text.replace(/\//g, "-").replace(" ", "T"));
  return Number.isFinite(parsed.getTime()) ? Math.floor(parsed.getTime() / 1000) : 0;
}

function indexedValue(data: Record<string, unknown>, keys: string[], index: number) {
  for (const key of keys) {
    const value = data[key];
    if (Array.isArray(value)) return value[index];
    if (value !== undefined) return value;
  }
  return undefined;
}

function extractPlayerInfoMap(payload: unknown, requestedIds: number[]) {
  const data = getApiData(payload);
  const map = new Map<number, Record<string, unknown>>();
  if (!data) return map;
  const list = getArray(data.List).length ? getArray(data.List) : getArray(data.Lst).length ? getArray(data.Lst) : getArray(data.Players).length ? getArray(data.Players) : getArray(data.PlayerInfo);
  if (list.length) {
    for (const item of list) {
      const player = getObject(item);
      if (!player) continue;
      const id = Number(player.UserId ?? player.Uid ?? player.Id ?? player.Pid);
      if (Number.isFinite(id)) map.set(id, player);
    }
    return map;
  }
  requestedIds.forEach((id, index) => {
    const player = {
      UserId: indexedValue(data, ["UserId", "Uid", "Id", "Pid"], index) ?? id,
      RegTime: indexedValue(data, ["RegTime", "RegisterTime", "Regt", "CreateTime"], index),
      Platform: indexedValue(data, ["Platform", "Plat", "System"], index),
      Version: indexedValue(data, ["Version", "AppVersion", "Ver"], index),
    };
    map.set(id, player);
  });
  return map;
}

function playerMatchesMailConditions(player: Record<string, unknown>, conditionRows: Array<{ field: string; op: string; value: string }>) {
  for (const row of conditionRows) {
    if (!row.value.trim()) continue;
    if (row.field === "regTime") {
      const regTime = parseTimeSeconds(player.RegTime ?? player.RegisterTime ?? player.Regt ?? player.CreateTime);
      const conditionTime = parseDatetimeLocalSeconds(dateToDatetimeLocal(row.value, row.op === "<="));
      if (!regTime || !conditionTime) return false;
      if (row.op === ">=" && regTime < conditionTime) return false;
      if (row.op === "<=" && regTime > conditionTime) return false;
    }
    if (row.field === "system") {
      const platforms = toPlatformNumberArray(row.value) as number[];
      const platform = Number(player.Platform ?? player.Plat ?? player.System);
      if (!platforms.length || !Number.isFinite(platform)) return false;
      if (row.op === "!=" ? platforms.includes(platform) : !platforms.includes(platform)) return false;
    }
    if (row.field === "version") {
      const versions = toVersionConditionArray([row]);
      const version = parseFlexibleNumber(String(player.Version ?? player.AppVersion ?? player.Ver ?? ""));
      if (!versions.length || !Number.isFinite(version)) return false;
      const matched = versions.includes(version);
      if (!matched) return false;
    }
  }
  return true;
}

function humanizeApiError(message: string) {
  if (/(?:Et|过期时间).*(?:当前时间|生效时间|发送时间).*(?:10\s*分钟|间隔|太短|过近)|(?:10\s*分钟|间隔太短).*(?:Et|过期时间)/i.test(message)) {
    return "过期时间间隔太短，请将过期时间设置为发送时间至少 11 分钟后";
  }
  const missingServerId = message.match(/(?:id|SvrId)\[(\d+)\].*(?:服务器不存在|不存在|未运行)/i);
  if (missingServerId) {
    const serverId = Number(missingServerId[1]);
    return `当前服务端未开放 ${gameServerDisplayName(serverId)}（真实ID ${serverId}），请确认区服是否已启动，或是否选错测试/正式服`;
  }
  if (/id\[\d+\].*服务器不存在/.test(message) || /服务器不存在/.test(message)) return "区服不存在，或选择的区服范围包含当前服务器未开放的区服";
  if (/TargetID|UserId|用户.*不存在|玩家.*不存在/i.test(message)) return "用户ID未填写，或用户不存在";
  if (/Platform|平台/i.test(message)) return "系统条件未填写，或选择的系统不支持";
  if (/Version|版本/i.test(message)) return "版本条件未填写，或版本格式不正确";
  if (/参数错误/.test(message)) return "条件参数填写不完整，请检查区服、系统、版本和时间";
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
    const response = await fetch(`/local-api/accounts/${accountId}`, { method: "DELETE" });
    if (!response.ok) {
      const payload = (await response.json().catch(() => ({}))) as { error?: string };
      throw new Error(payload.error || "删除账号失败");
    }
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
    const response = await fetch(`/local-api/games/${gameId}`, { method: "DELETE" });
    if (!response.ok) {
      const payload = (await response.json().catch(() => ({}))) as { error?: string };
      throw new Error(payload.error || "删除区服失败");
    }
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

  const sharedPanels = (
    <>
      {auth.isAdmin && accountPanelOpen && <AccountPanel accounts={accounts} games={games} canManageAdmins={auth.isRootAdmin} session={session} onAdd={createAccount} onDelete={deleteAccount} onUpdate={updateAccount} onClose={() => setAccountPanelOpen(false)} />}
      {auth.isAdmin && gamePanelOpen && <GamePanel games={games} onAdd={createGame} onDelete={deleteGame} onUpdate={updateGame} onClose={() => setGamePanelOpen(false)} />}
      {auth.isAdmin && logPanelOpen && <UserLogPanel onClose={() => setLogPanelOpen(false)} />}
      {profilePanelOpen && auth && <ProfilePanel auth={auth} onClose={() => setProfilePanelOpen(false)} onSave={updateProfile} />}
    </>
  );

  if (!session) {
    return (
      <>
        <GameSelectScreen
          auth={auth}
          games={games}
          onEnter={setSession}
          onLogout={logout}
          onManageAccounts={() => setAccountPanelOpen(true)}
          onManageGames={() => setGamePanelOpen(true)}
          onOpenLogs={() => setLogPanelOpen(true)}
          onOpenProfile={() => setProfilePanelOpen(true)}
          onReorder={reorderGames}
        />
        {sharedPanels}
        {versionModal}
      </>
    );
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

          {active === "dashboard" ? <Dashboard results={results} accounts={accounts} /> : active === "playerInfo" ? <PlayerInfoPage postWithToken={postWithToken} /> : active === "bindUid" ? <BindUidPage postWithToken={postWithToken} /> : active === "gmState" ? <BanControlPage postWithToken={postWithToken} /> : active === "silence" ? <SilencePage operator={session.operatorAccount} /> : active.startsWith("mail") ? <MailSuitePage active={active as MailSectionKey} canUploadItemTable={auth.isAdmin} postWithToken={postWithToken} session={session} setActive={setActive} /> : active.startsWith("gift") ? <GiftSuitePage active={active as GiftSectionKey} canUploadItemTable={auth.isAdmin} postWithToken={postWithToken} /> : active === "notice" ? <NoticePage postWithToken={postWithToken} /> : active === "noticeTemplate" ? <NoticeTemplatePage /> : active === "serverTime" ? <OpenServerPage postWithToken={postWithToken} /> : active === "orderRefund" ? <OrderRefundPage postWithToken={postWithToken} /> : moduleConfig.status === "pending" ? <UnavailablePanel module={moduleConfig} /> : (
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
      {sharedPanels}
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
            <label>原密码<input autoComplete="current-password" type="password" value={oldPassword} onChange={(event) => setOldPassword(event.target.value)} /></label>
            <label>新密码<input autoComplete="new-password" type="password" value={newPassword} onChange={(event) => setNewPassword(event.target.value)} /></label>
            <label>确认新密码<input autoComplete="new-password" type="password" value={confirmPassword} onChange={(event) => setConfirmPassword(event.target.value)} /></label>
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

function GameSelectScreen({ auth, games, onEnter, onLogout, onManageAccounts, onManageGames, onOpenLogs, onOpenProfile, onReorder }: { auth: AuthSession; games: GameConfig[]; onEnter: (session: Session) => void; onLogout: () => void; onManageAccounts: () => void; onManageGames: () => void; onOpenLogs: () => void; onOpenProfile: () => void; onReorder: (orderedIds: number[]) => Promise<void> }) {
  const [error, setError] = React.useState("");
  const [entering, setEntering] = React.useState("");
  const [orderedGames, setOrderedGames] = React.useState<GameConfig[]>([]);
  const [draggingName, setDraggingName] = React.useState<string | null>(null);
  const [pressingName, setPressingName] = React.useState<string | null>(null);
  const allowedGames = React.useMemo(() => games.filter((game) => auth.isAdmin || auth.games.includes(`${game.name}/${game.serverName}`)), [auth.games, auth.isAdmin, games]);
  const allowedSignature = React.useMemo(() => allowedGames.map((game) => game.id ?? `${game.name}/${game.serverName}`).join("|"), [allowedGames]);
  const pressTimerRef = React.useRef<number | null>(null);
  const draggedNameRef = React.useRef<string | null>(null);
  const dragMovedRef = React.useRef(false);
  const suppressClickRef = React.useRef(false);
  const latestOrderRef = React.useRef<GameConfig[]>(allowedGames);
  const gameGroups = React.useMemo(() => {
    const groups: Array<{ name: string; games: GameConfig[] }> = [];
    for (const game of orderedGames) {
      const group = groups.find((item) => item.name === game.name);
      if (group) group.games.push(game);
      else groups.push({ name: game.name, games: [game] });
    }
    return groups;
  }, [orderedGames]);

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

  const beginPress = (event: React.PointerEvent<HTMLElement>, gameName: string) => {
    if (entering) return;
    const card = event.currentTarget;
    const pointerId = event.pointerId;
    clearPressTimer();
    setPressingName(gameName);
    pressTimerRef.current = window.setTimeout(() => {
      draggedNameRef.current = gameName;
      dragMovedRef.current = false;
      suppressClickRef.current = true;
      setDraggingName(gameName);
      setPressingName(null);
      card.setPointerCapture(pointerId);
    }, 420);
  };

  const moveDraggedGameGroup = (targetName: string) => {
    const sourceName = draggedNameRef.current;
    if (!sourceName || sourceName === targetName) return;
    setOrderedGames((current) => {
      const currentGroups: Array<{ name: string; games: GameConfig[] }> = [];
      for (const game of current) {
        const group = currentGroups.find((item) => item.name === game.name);
        if (group) group.games.push(game);
        else currentGroups.push({ name: game.name, games: [game] });
      }
      const fromIndex = currentGroups.findIndex((group) => group.name === sourceName);
      const toIndex = currentGroups.findIndex((group) => group.name === targetName);
      if (fromIndex < 0 || toIndex < 0 || fromIndex === toIndex) return current;
      const nextGroups = [...currentGroups];
      const [moved] = nextGroups.splice(fromIndex, 1);
      nextGroups.splice(toIndex, 0, moved);
      const next = nextGroups.flatMap((group) => group.games);
      latestOrderRef.current = next;
      dragMovedRef.current = true;
      return next;
    });
  };

  const continueDrag = (event: React.PointerEvent<HTMLElement>) => {
    if (!draggedNameRef.current) return;
    event.preventDefault();
    const target = document.elementFromPoint(event.clientX, event.clientY)?.closest("[data-game-name]");
    const targetName = target?.getAttribute("data-game-name") ?? "";
    if (targetName) moveDraggedGameGroup(targetName);
  };

  const finishDrag = async (event?: React.PointerEvent<HTMLElement>) => {
    clearPressTimer();
    setPressingName(null);
    const draggedName = draggedNameRef.current;
    draggedNameRef.current = null;
    setDraggingName(null);
    try {
      if (event?.currentTarget.hasPointerCapture(event.pointerId)) {
        event.currentTarget.releasePointerCapture(event.pointerId);
      }
    } catch {
      // Some browsers throw if capture was already released.
    }
    if (!draggedName || !dragMovedRef.current) {
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
            {auth.isAdmin && (
              <>
                <button className="account-button" onClick={onManageGames} type="button"><Plus size={15} />游戏区服</button>
                <button className="account-button" onClick={onManageAccounts} type="button"><UserCog size={15} />账号管理</button>
                <button className="account-button" onClick={onOpenLogs} type="button"><Database size={15} />用户日志</button>
              </>
            )}
            <button className="profile-button" onClick={onOpenProfile} type="button" title="个人信息设置">
              <AvatarImage account={auth.operatorAccount} avatarUrl={auth.avatarUrl} displayName={auth.displayName || auth.operatorAccount} />
              <span>{auth.displayName || auth.operatorAccount}</span>
            </button>
            <button className="icon-button" onClick={onLogout} type="button" title="退出"><LogOut size={18} /></button>
          </div>
        </header>
        <div className="game-select-content">
          <h1>游戏列表</h1>
          {allowedGames.length === 0 ? (
            <section className="empty-game-list"><ShieldBan size={30} /><strong>暂无可访问游戏</strong><span>请联系管理员为该账号分配游戏区服权限。</span></section>
          ) : (
            <section className={`game-card-grid${draggingName ? " drag-active" : ""}`}>
              {gameGroups.map((group) => {
                const primaryGame = group.games[0];
                const isDragging = draggingName === group.name;
                const isPressing = pressingName === group.name;
                return (
                  <article
                    className={`game-card${isDragging ? " dragging" : ""}${isPressing ? " press-ready" : ""}`}
                    data-game-name={group.name}
                    key={group.name}
                    onPointerCancel={(event) => void finishDrag(event)}
                    onPointerDown={(event) => beginPress(event, group.name)}
                    onPointerLeave={clearPressTimer}
                    onPointerMove={continueDrag}
                    onPointerUp={(event) => void finishDrag(event)}
                  >
                    <div className="game-cover" style={primaryGame.backgroundUrl ? { backgroundImage: `url(${primaryGame.backgroundUrl})` } : undefined}>
                      {!primaryGame.backgroundUrl && <span>{group.name}</span>}
                    </div>
                    <div className="game-card-body">
                      <div className="game-logo">{primaryGame.iconUrl ? <img alt={`${group.name} icon`} src={primaryGame.iconUrl} /> : primaryGame.logo}</div>
                      <div className="game-meta">
                        <strong>{group.name}</strong>
                        <small>{group.games.length} 个区服</small>
                        <div className="game-server-links">
                          {group.games.map((game) => {
                            const key = `${game.name}/${game.serverName}`;
                            return (
                              <button disabled={entering === key} key={game.id ?? key} onClick={() => void enterGame(game)} type="button">
                                {entering === key ? "进入中..." : game.serverName}
                              </button>
                            );
                          })}
                        </div>
                      </div>
                    </div>
                  </article>
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
      const result = await postWithToken("/gmQueryPlayerInfo", { UserId: ids[0] });
      const parsed = result.payload as Record<string, unknown>;
      setPayload(parsed);
      const businessError = apiBusinessError(result);
      if (businessError) setError(businessError === "用户ID未填写，或用户不存在" ? "用户id不存在" : businessError);
    } catch (error) {
      setError(error instanceof Error ? error.message : "查询失败");
      setPayload(null);
    } finally {
      setLoading(false);
    }
  };

  const data = getObject(payload?.data) ?? getObject(payload);
  const info = getObject(data?.Info);
  const extra = getObject(data?.Extra)
    ?? getObject(data?.extra)
    ?? getObject(data?.Info2)
    ?? getObject(data?.info2)
    ?? getObject(data?.GmExtra)
    ?? getObject(data?.gmExtra)
    ?? getObject(data?.PlayerInfoExtra)
    ?? getObject(data?.playerInfoExtra)
    ?? getObject(info?.Extra)
    ?? getObject(info?.extra)
    ?? getObject(info?.Info2)
    ?? getObject(info?.info2);
  const extraValue = (...keys: string[]) => {
    for (const key of keys) {
      const value = extra?.[key] ?? data?.[key] ?? info?.[key];
      if (value !== undefined && value !== null && value !== "") return value;
    }
    return undefined;
  };
  const dataValue = (...keys: string[]) => {
    for (const key of keys) {
      const value = data?.[key] ?? info?.[key] ?? extra?.[key];
      if (value !== undefined && value !== null && value !== "") return value;
    }
    return undefined;
  };
  const languageValue = extraValue("language", "Language");
  const languageName = languageDefinitions.find((language) => language.id === Number(languageValue))?.label ?? languageValue;
  const platformValue = extraValue("platform", "Platform");
  const platformName = Number(platformValue) === 1 ? "GooglePlay" : Number(platformValue) === 2 ? "iOS" : platformValue;
  const clientVersionValue = extraValue("clientVersion", "ClientVersion", "client_version");
  const onlineValue = extraValue("online", "Online");
  const firstUserId = dataValue("UserId", "UID", "uid", "Uid") ?? filters.userId;
  const desc = typeof data?.Desc === "string" ? data.Desc : "";
  const mapAttribValue = getObject(info?.mapAttribValue);
  const mapAttribBase = getObject(info?.mapAttribBase);
  const mapAttribRatio = getObject(info?.mapAttribRatio);
  const mapUnitLvl = getObject(info?.mapUnitLvl);
  const formations = getArray(info?.lstFormation).map((formation, index) => {
    const formationObject = getObject(formation);
    return {
      index: `阵容${index + 1}`,
      troopIds: getArray(formationObject?.troopIds).join(", "),
    };
  });
  const gears = Object.entries(getObject(info?.mapGear) ?? {}).map(([site, item]) => {
    const itemObject = getObject(item);
    return {
      id: itemObject?.id ?? site,
      name: itemObject?.id ?? site,
      level: formatCell(itemObject?.lvl),
      using: "是",
    };
  });
  const heroes = Object.entries(mapUnitLvl ?? {}).map(([id, level]) => ({
    id,
    name: id,
    star: formatCell(level),
    using: Number(info?.heroId) === Number(id) || Number(info?.idHeroUsing) === Number(id) ? "是" : "否",
  }));
  const runeRows = getArray(info?.inlayInfo).map((item, index) => ({
    id: `镶嵌${index + 1}`,
    name: JSON.stringify(getObject(item)?.dict ?? {}),
    using: "是",
  }));
  const orderRows = [
    { label: "宠物", value: JSON.stringify(info?.pet ?? {}) },
    { label: "宝石背包", value: JSON.stringify(info?.gemBag ?? {}) },
    { label: "基础属性", value: JSON.stringify(mapAttribBase ?? {}) },
    { label: "属性倍率", value: JSON.stringify(mapAttribRatio ?? {}) },
  ].filter((row) => row.value && row.value !== "{}");
  const summaryRows = data
    ? [{
      userPid: firstUserId,
      deviceId: dataValue("DeviceId", "deviceId"),
      platformUid: dataValue("AccId", "AccountId", "platformUid", "PlatformUid"),
      registerTime: formatTimestampValue(extraValue("reg", "Reg", "RegTime", "regTime")),
      lastLoginTime: formatTimestampValue(extraValue("out", "Out", "LastLoginTime", "lastLoginTime")),
    }]
    : [];
  const infoCells: Array<[string, unknown]> = [
    ["用户PID", firstUserId],
    ["deviceId", dataValue("DeviceId", "deviceId")],
    ["platformUid", dataValue("AccId", "AccountId", "platformUid", "PlatformUid")],
    ["金币", extraValue("goldNum", "GoldNum") ?? mapAttribValue?.["4"] ?? data?.Gold],
    ["钻石", extraValue("diamondNum", "DiamondNum") ?? data?.Diamond],
    ["点券", extraValue("couponNum", "CouponNum")],
    ["退款翡翠欠款", extraValue("refundJadeDebt", "RefundJadeDebt")],
    ["退款点券欠款", extraValue("refundItem999Debt", "RefundItem999Debt")],
    ["等级", info?.lvl ?? info?.grade ?? data?.Level],
    ["注册时间", formatTimestampValue(extraValue("reg", "Reg", "RegTime", "regTime"))],
    ["最后登出时间", formatTimestampValue(extraValue("out", "Out", "LastLoginTime", "lastLoginTime"))],
    ["国家", dataValue("Country", "country")],
    ["状态", desc || data?.State],
    ["在线状态", typeof onlineValue === "boolean" ? onlineValue ? "在线" : "离线" : onlineValue],
    ["是否付费", dataValue("IsPay", "isPay")],
    ["总充值金额", extraValue("cntCharge", "CntCharge") ?? data?.PayTotal],
    ["总消耗钻石数量", dataValue("CostDiamond", "costDiamond")],
    ["总消耗金币数量", dataValue("CostGold", "costGold")],
    ["关卡ID", info?.maxChapter ?? data?.StageId],
    ["系统标识", platformName ?? data?.System],
    ["玩家语言", languageName],
    ["客户端通信版本号", clientVersionValue !== undefined ? versionCodeToText(clientVersionValue) : data?.ClientVersion],
    ["头像ID", info?.head],
    ["头像框ID", info?.headFrameId],
    ["称号ID", info?.titleId],
    ["战力", info?.combatPower],
    ["当前地图ID", info?.mapid],
    ["当前英雄ID", info?.heroId ?? info?.idHeroUsing],
    ["竞技场排名", Number(info?.arenaRank) === 4294967295 ? "未上榜" : info?.arenaRank],
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
        {activeTab === "用户信息" ? <PlayerInfoGrid cells={infoCells} /> : activeTab === "用户行为" && formations.length ? <PlayerSimpleTable columns={["阵容", "上阵单位"]} keys={["index", "troopIds"]} emptyText="暂无数据" rows={formations} /> : activeTab === "用户订单" && orderRows.length ? <PlayerSimpleTable columns={["项目", "内容"]} keys={["label", "value"]} emptyText="暂无数据" rows={orderRows} /> : <div className="player-empty-block">当前服务器接口暂未开放“{activeTab}”数据</div>}
      </section>

      <PlayerSectionTable title="游戏角色" columns={["角色ID", "角色名称", "星级", "正在使用"]} keys={["id", "name", "star", "using"]} rows={heroes} />
      <PlayerSectionTable title="角色装备" columns={["装备ID", "装备名称", "等级", "正在使用"]} keys={["id", "name", "level", "using"]} rows={gears} />
      <PlayerSectionTable title="角色符文" columns={["符文ID", "符文名称", "正在使用"]} keys={["id", "name", "using"]} rows={runeRows} />
    </section>
  );
}

function PlayerSearchField({ disabled, label, onChange, placeholder, value }: { disabled?: boolean; label: string; onChange: (value: string) => void; placeholder: string; value: string }) {
  return <label className="player-search-field"><span>{label}:</span><input disabled={disabled} value={value} onChange={(event) => onChange(event.target.value)} placeholder={placeholder} /></label>;
}

function PlayerSimpleTable({ columns, emptyText, keys = ["userPid", "deviceId", "platformUid", "registerTime", "lastLoginTime"], rows }: { columns: string[]; emptyText: string; keys?: string[]; rows: Array<Record<string, unknown>> }) {
  return (
    <section className="player-table">
      <table><thead><tr>{columns.map((column) => <th key={column}>{column}</th>)}</tr></thead><tbody>{rows.length ? rows.map((row, index) => <tr key={index}>{keys.map((key) => <td key={key}>{formatCell(row[key])}</td>)}</tr>) : <tr><td colSpan={columns.length}>{emptyText}</td></tr>}</tbody></table>
    </section>
  );
}

function PlayerInfoGrid({ cells }: { cells: Array<[string, unknown]> }) {
  return <div className="player-info-grid">{cells.map(([label, value]) => <React.Fragment key={label}><div className="cell-label">{label}</div><div>{formatCell(value)}</div></React.Fragment>)}</div>;
}

function PlayerSectionTable({ columns, keys, rows, title }: { columns: string[]; keys: string[]; rows: Array<Record<string, unknown>>; title: string }) {
  return <section className="player-section-table"><div>{title}</div><PlayerSimpleTable columns={columns} emptyText="暂无数据" keys={keys} rows={rows} /></section>;
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

function OrderRefundPage({ postWithToken }: { postWithToken: (endpoint: string, body: unknown) => Promise<ApiPostResponse> }) {
  const [orderId, setOrderId] = React.useState("");
  const [status, setStatus] = React.useState("");
  const [submitting, setSubmitting] = React.useState(false);
  const [lastResult, setLastResult] = React.useState<Record<string, unknown> | null>(null);

  const submit = async () => {
    const normalizedOrderId = orderId.trim();
    if (!normalizedOrderId) {
      setStatus("请输入订单号");
      return;
    }
    if (!window.confirm(`确认对订单 ${normalizedOrderId} 发起退款？`)) return;
    setSubmitting(true);
    setStatus("");
    setLastResult(null);
    try {
      const result = await postWithToken("/gmOrderRefund", { order_id: normalizedOrderId });
      const payload = getObject(result.payload) ?? {};
      const data = getApiData(result.payload);
      const code = Number(payload.code ?? data?.code);
      const message = String(payload.result ?? data?.result ?? "");
      setLastResult(payload);
      if (!result.ok) {
        setStatus(`退款失败：HTTP ${result.status}`);
        return;
      }
      if (Number.isFinite(code) && code !== 0) {
        setStatus(`退款失败：${message || `错误码 ${code}`}`);
        return;
      }
      setStatus(`退款成功：${normalizedOrderId}`);
    } catch (error) {
      setStatus(error instanceof Error ? `退款失败：${error.message}` : "退款失败");
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <section className="order-refund-page">
      <div className="refund-card">
        <header><CreditCard size={22} /><div><strong>三方支付用户退款</strong><span>输入三方支付订单号后提交给游戏服务端处理退款。</span></div></header>
        <label>订单号<input value={orderId} onChange={(event) => setOrderId(event.target.value)} placeholder="例如：GPA.3340-7674-3284-test" /></label>
        <div className="refund-actions"><button disabled={submitting} onClick={() => void submit()} type="button">{submitting ? "处理中..." : "处理退款"}</button></div>
        {status && <div className={status.includes("成功") ? "refund-status success" : "refund-status"}>{status}</div>}
        {lastResult && <pre className="refund-result">{JSON.stringify(lastResult, null, 2)}</pre>}
      </div>
    </section>
  );
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
    const [itemResponse, templateResponse, rewardResponse, serverResponse, scheduledResponse, groupResponse] = await Promise.all([
      fetch("/local-api/items"),
      fetch("/local-api/mail-templates"),
      fetch("/local-api/reward-templates"),
      fetch("/local-api/game-servers"),
      fetch(`/local-api/scheduled-mails?serverUrl=${encodeURIComponent(session.serverUrl)}`),
      fetch(`/local-api/mail-groups?serverUrl=${encodeURIComponent(session.serverUrl)}`),
    ]);
    const itemPayload = (await itemResponse.json().catch(() => ({}))) as { items?: ItemOption[] };
    const templatePayload = (await templateResponse.json().catch(() => ({}))) as { templates?: MailTemplate[] };
    const rewardPayload = (await rewardResponse.json().catch(() => ({}))) as { templates?: RewardTemplate[] };
    const serverPayload = (await serverResponse.json().catch(() => ({}))) as { servers?: ServerOption[] };
    const scheduledPayload = (await scheduledResponse.json().catch(() => ({}))) as { mails?: Record<string, unknown>[] };
    const groupPayload = (await groupResponse.json().catch(() => ({}))) as { groups?: Record<string, unknown>[] };
    setItems(itemPayload.items ?? []);
    setTemplates(templatePayload.templates ?? []);
    setRewardTemplates(rewardPayload.templates ?? []);
    setServerOptions(serverPayload.servers ?? []);
    setLocalMailRows([...(groupPayload.groups ?? []), ...(scheduledPayload.mails ?? [])]);
  }, [session.serverUrl]);

  const persistMailGroup = async (group: Record<string, unknown>) => {
    await fetch("/local-api/mail-groups", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ serverUrl: session.serverUrl, group }),
    }).catch(() => undefined);
  };

  const removeMailGroup = async (id: string) => {
    await fetch(`/local-api/mail-groups/${encodeURIComponent(id)}`, { method: "DELETE" }).catch(() => undefined);
  };

  const refreshMailList = React.useCallback(async () => {
    const result = await postWithToken("/gmMailLst", {});
    const error = apiBusinessError(result);
    if (error) {
      setStatus(`邮件列表读取失败：${error}`);
      setMailRows([]);
      return;
    }
    const data = getApiData(result.payload);
    const globalRows = getArray(data?.Lst)
      .map((row) => getObject(row))
      .filter((row): row is Record<string, unknown> => Boolean(row))
      .map((row) => ({ ...row, __mailListType: "global" }));
    const personalRows = getArray(data?.LstUser)
      .map((row) => getObject(row))
      .filter((row): row is Record<string, unknown> => Boolean(row))
      .map((row) => ({ ...row, Typ: row.Typ ?? 3, __mailListType: "personal" }));
    const rows = [...globalRows, ...personalRows];
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
  const recallEditingMail = async () => {
    if (!editingMailRow) return;
    const id = String(editingMailRow.Id ?? editingMailRow.id ?? "");
    if (!id) return;
    if (editingMailRow.__scheduled) {
      const response = await fetch(`/local-api/scheduled-mails/${encodeURIComponent(id)}`, { method: "DELETE" });
      const payload = (await response.json().catch(() => ({}))) as { error?: string };
      if (!response.ok) throw new Error(payload.error || "原定时邮件撤回失败");
      setLocalMailRows((current) => current.filter((row) => String(row.Id) !== id));
      return;
    }
    const groupIds = mailGroupIds(editingMailRow);
    const failures: string[] = [];
    for (const mailId of groupIds.length ? groupIds : [id]) {
      const result = await postWithToken("/gmMailDel", { Id: mailId });
      const error = apiBusinessError(result);
      if (error) failures.push(`${mailId}：${error}`);
    }
    if (failures.length) throw new Error(`原邮件撤回失败：${failures[0]}`);
    if (groupIds.length > 1) await removeMailGroup(id);
    setLocalMailRows((current) => current.filter((row) => String(row.Id) !== id));
  };

  if (active === "mailTemplate") {
    return view === "edit"
      ? <MailTemplateEditor onBack={() => setView("list")} onSaved={(savedTemplate) => {
        if (savedTemplate) setTemplates((current) => [savedTemplate, ...current.filter((template) => template.id !== savedTemplate.id)]);
        setView("list");
        setEditingMailTemplate(undefined);
        void refreshLocalMailData();
      }} template={editingMailTemplate} />
      : <MailTemplateList query={templateQuery} setQuery={setTemplateQuery} templates={templates} onCreate={() => { setEditingMailTemplate(undefined); setView("edit"); }} onEdit={(template) => { setEditingMailTemplate(template); setView("edit"); }} onDelete={async (template) => {
        if (!window.confirm(`确认删除邮件模板「${template.name}」？`)) return;
        await fetch(`/local-api/mail-templates/${encodeURIComponent(template.id)}`, { method: "DELETE" });
        await refreshLocalMailData();
      }} />;
  }

  if (active === "mailRewardTemplate") {
    return view === "edit"
      ? <RewardTemplateEditor canUploadItemTable={canUploadItemTable} items={items} onBack={() => setView("list")} onSaved={(savedTemplate) => {
        setRewardTemplates((current) => [savedTemplate, ...current.filter((item) => item.id !== savedTemplate.id)]);
        setView("list");
        setEditingRewardTemplate(undefined);
        void refreshLocalMailData();
      }} onUploadItemTable={uploadItemTable} template={editingRewardTemplate} />
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
          const conditionRows = getArray(submitted.__conditionRows)
            .map((row) => getObject(row))
            .filter((row): row is { field: string; op: string; value: string } => Boolean(row))
            .map((row) => ({ field: String(row.field ?? ""), op: String(row.op ?? "="), value: String(row.value ?? "") }));
          const serverBody: Record<string, unknown> = { ...submitted };
          delete serverBody.__conditionRows;
          const isPersonalMailSubmit = Number(serverBody.Typ) === 3;
          const submitNowSeconds = Math.floor(Date.now() / 1000);
          const submitStartSeconds = Number(serverBody.St);
          const submitEndSeconds = Number(serverBody.Et);
          const effectiveStartSeconds = Number.isFinite(submitStartSeconds) && submitStartSeconds > submitNowSeconds
            ? submitStartSeconds
            : submitNowSeconds;
          if (!Number.isFinite(submitEndSeconds) || submitEndSeconds <= effectiveStartSeconds + 10 * 60) {
            throw new Error("过期时间间隔太短，请将过期时间设置为发送时间至少 11 分钟后");
          }
          if (isPersonalMailSubmit) {
            const targets = getArray(serverBody.TargetID).map((item) => Number(item)).filter((item) => Number.isFinite(item));
            const infoResult = await postWithToken("/gmPlayerInfo", { Typ: 1, UserId: targets });
            const infoError = apiBusinessError(infoResult);
            if (infoError) throw new Error("用户id不存在");
            const playerMap = extractPlayerInfoMap(infoResult.payload, targets);
            if (targets.some((targetId) => !playerMap.has(targetId))) {
              throw new Error("用户id不存在");
            }
            if (!conditionRows.some((row) => row.value.trim())) {
              serverBody.RegtBegin = 0;
              serverBody.Regt = 0;
              serverBody.Platform = [];
              serverBody.Version = [];
            } else {
            const matchedTargets = targets.filter((targetId) => {
              const player = playerMap.get(targetId);
              return player ? playerMatchesMailConditions(player, conditionRows) : false;
            });
            if (!matchedTargets.length) {
              throw new Error("没有符合条件的用户，邮件未发送");
            }
            if (matchedTargets.length < targets.length) {
              const removed = targets.filter((targetId) => !matchedTargets.includes(targetId));
              setStatus(`已按条件过滤，跳过 ${removed.length} 个不符合条件的用户`);
            }
            serverBody.TargetID = matchedTargets;
            serverBody.RegtBegin = 0;
            serverBody.Regt = parseDatetimeLocalSeconds(defaultRegEndTime());
            serverBody.Platform = [];
            serverBody.Version = [];
            }
          }
          const startSeconds = Number(submitted.St);
          const isScheduled = Number.isFinite(startSeconds) && startSeconds > Math.floor(Date.now() / 1000) + 5;
          if (editingMailRow) {
            setStatus("正在撤回原邮件...");
            await recallEditingMail();
          }
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
                body: serverBody,
              }),
            });
            const payload = (await response.json().catch(() => ({}))) as { mail?: Record<string, unknown>; error?: string };
            if (!response.ok || !payload.mail) throw new Error(payload.error || "创建定时邮件失败");
            const localRow: Record<string, unknown> = { ...(payload.mail ?? {}), __conditionRows: conditionRows };
            const message = `邮件已加入定时队列，将在 ${formatTimestamp(startSeconds)} 发送`;
            setStatus(message);
            setView("list");
            setEditingMailRow(undefined);
            setLocalMailRows((current) => [localRow, ...current.filter((row) => String(row.Id) !== String(localRow.Id))].slice(0, 50));
            return message;
          }
          const versionTargets = getArray(serverBody.Version).map((item) => Number(item)).filter((item) => Number.isFinite(item));
          const versionChunks = versionTargets.length > 2000 ? chunkArray(versionTargets, 2000) : [];
          if (Number(serverBody.Typ) === 1 && versionChunks.length > 1) {
            const successes: Array<{ chunk: number[]; data: Record<string, unknown> | null }> = [];
            const failures: string[] = [];
            setStatus(`正在按版本分片发送：0/${versionChunks.length}`);
            for (const [index, versionChunk] of versionChunks.entries()) {
              try {
                const result = await postWithToken("/gmMailAdd", { ...serverBody, Version: versionChunk });
                const error = apiBusinessError(result);
                if (error) {
                  failures.push(error);
                } else {
                  successes.push({ chunk: versionChunk, data: getApiData(result.payload) });
                }
              } catch (sendError) {
                failures.push(sendError instanceof Error ? sendError.message : "发送失败");
              }
              setStatus(`正在按版本分片发送：${index + 1}/${versionChunks.length}，成功 ${successes.length}，失败 ${failures.length}`);
            }
            if (!successes.length) {
              const firstError = failures[0] || "没有版本分片发送成功";
              throw new Error(`邮件提交失败：${firstError}`);
            }
            const firstData = successes[0]?.data ?? null;
            const localRow = {
              ...submitted,
              ...(firstData ?? {}),
              Id: String(firstData?.Id ?? submitted.Id ?? `local-${Date.now()}`),
              Version: versionTargets,
              RegtEnd: firstData?.RegtEnd ?? submitted.Regt,
              CreateTime: firstData?.CreateTime ?? Math.floor(Date.now() / 1000),
              __targetSummary: `全服，版本分片 ${successes.length}/${versionChunks.length} 成功`,
              __childMailIds: successes.map((item) => item.data?.Id ?? item.data?.id).filter((item) => item !== undefined && item !== null && String(item).trim()),
              __conditionRows: conditionRows,
              __local: !firstData?.Id,
              __claimed: false,
            };
            const message = failures.length
              ? `邮件已按版本分片发送，成功 ${successes.length} 包，失败 ${failures.length} 包`
              : `邮件已按版本分片发送，成功 ${successes.length} 包`;
            await persistMailGroup(localRow);
            setStatus(message);
            setView("list");
            setEditingMailRow(undefined);
            setLocalMailRows((current) => [localRow, ...current.filter((row) => String(row.Id) !== String(localRow.Id))].slice(0, 20));
            await refreshMailList();
            if (active !== "mailGlobal") setActive("mailGlobal");
            setStatus(message);
            return message;
          }
          const serverMailTargets = Number(serverBody.Typ) === 2 ? getArray(serverBody.TargetID).map((item) => Number(item)).filter((item) => Number.isFinite(item)) : [];
          if (serverMailTargets.length > 1) {
            const successes: Array<{ targetId: number; data: Record<string, unknown> | null }> = [];
            const failures: Array<{ targetId: number; error: string }> = [];
            let cursor = 0;
            let completed = 0;
            const concurrency = 8;
            setStatus(`正在发送区服邮件：0/${serverMailTargets.length}`);
            const sendOne = async (targetId: number) => {
              try {
                const result = await postWithToken("/gmMailAdd", { ...serverBody, TargetID: [targetId] });
                const error = apiBusinessError(result);
                if (error) {
                  failures.push({ targetId, error });
                } else {
                  successes.push({ targetId, data: getApiData(result.payload) });
                }
              } catch (sendError) {
                failures.push({ targetId, error: sendError instanceof Error ? sendError.message : "发送失败" });
              } finally {
                completed += 1;
                setStatus(`正在发送区服邮件：${completed}/${serverMailTargets.length}，成功 ${successes.length}，失败 ${failures.length}`);
              }
            };
            const workers = Array.from({ length: Math.min(concurrency, serverMailTargets.length) }, async () => {
              while (cursor < serverMailTargets.length) {
                const targetId = serverMailTargets[cursor];
                cursor += 1;
                await sendOne(targetId);
              }
            });
            await Promise.all(workers);
            if (!successes.length) {
              const firstError = failures[0]?.error || "没有区服发送成功";
              throw new Error(`邮件提交失败：${firstError}`);
            }
            const firstData = successes[0]?.data ?? null;
            const successRange = formatServerRange(successes.map((item) => item.targetId));
            const localRow = {
              ...submitted,
              ...(firstData ?? {}),
              Id: String(firstData?.Id ?? submitted.Id ?? `local-${Date.now()}`),
              TargetID: successes.map((item) => item.targetId),
              RegtEnd: firstData?.RegtEnd ?? submitted.Regt,
              CreateTime: firstData?.CreateTime ?? Math.floor(Date.now() / 1000),
              __targetSummary: `${successRange || "指定区服"} 成功`,
              __childMailIds: successes.map((item) => item.data?.Id ?? item.data?.id).filter((item) => item !== undefined && item !== null && String(item).trim()),
              __conditionRows: conditionRows,
              __local: !firstData?.Id,
              __claimed: false,
            };
            const message = failures.length
              ? `邮件已发送到 ${successes.length} 个区服，${failures.length} 个区服未发送`
              : `邮件已发送到 ${successes.length} 个区服`;
            await persistMailGroup(localRow);
            setStatus(message);
            setView("list");
            setEditingMailRow(undefined);
            setLocalMailRows((current) => [localRow, ...current.filter((row) => String(row.Id) !== String(localRow.Id))].slice(0, 20));
            await refreshMailList();
            if (active !== "mailGlobal") setActive("mailGlobal");
            setStatus(message);
            return message;
          }
          const result = await postWithToken("/gmMailAdd", serverBody);
          const error = apiBusinessError(result);
          if (error) {
            const message = isPersonalMailSubmit && /条件参数|用户ID|用户不存在|TargetID/i.test(error) ? "用户id不存在" : error;
            throw new Error(`邮件提交失败：${message}`);
          }
          const data = getApiData(result.payload);
          const localRow = {
            ...submitted,
            ...(data ?? {}),
            Id: String(data?.Id ?? submitted.Id ?? `local-${Date.now()}`),
            RegtEnd: data?.RegtEnd ?? submitted.Regt,
            CreateTime: data?.CreateTime ?? Math.floor(Date.now() / 1000),
            __conditionRows: conditionRows,
            __local: !data?.Id,
            __claimed: false,
          };
          const message = `邮件已提交到服务器${data?.Id ? `，ID：${String(data.Id)}` : ""}`;
          setStatus(message);
          const submittedTyp = Number(serverBody.Typ);
          const submittedTargets = getArray(serverBody.TargetID);
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
      onClearStatus={() => setStatus("")}
      userIdQuery={userIdQuery}
      setUserIdQuery={setUserIdQuery}
      onCreate={() => { setEditingMailRow(undefined); setView("edit"); }}
      onEdit={async (row) => {
        const id = String(row.Id ?? row.id ?? "");
        const typ = Number(row.Typ);
        const isPersonalMail = row.__mailListType === "personal" || typ === 3 || (!typ && getArray(row.TargetID).length > 0);
        if (isPersonalMail) {
          let latestRow = row;
          try {
            const result = await postWithToken("/gmMailLst", {});
            const data = getApiData(result.payload);
            latestRow = getArray(data?.LstUser)
              .map((item) => getObject(item))
              .filter((item): item is Record<string, unknown> => Boolean(item))
              .find((item) => String(item.Id ?? item.id ?? "") === id) ?? row;
          } catch {
            latestRow = row;
          }
          const hasRewards = getArray(latestRow.ItemLst).length > 0;
          if (mailRowClaimed(latestRow) || (hasRewards && !mailRowClaimStateKnown(latestRow))) {
            setStatus("邮件被领取，不能进行编辑重新发送");
            return;
          }
        }
        setEditingMailRow(row);
        setView("edit");
      }}
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
        const groupIds = mailGroupIds(localRow);
        const deleteIds = groupIds.length ? groupIds : [id];
        const failures: string[] = [];
        for (const mailId of deleteIds) {
          const result = await postWithToken("/gmMailDel", { Id: mailId });
          const error = apiBusinessError(result);
          if (error) failures.push(error);
        }
        const successCount = deleteIds.length - failures.length;
        setStatus(failures.length ? successCount ? `已撤回 ${successCount} 封，${failures.length} 封撤回失败` : `删除失败：${failures[0]}` : "邮件已删除");
        if (successCount > 0) {
          if (groupIds.length > 1) await removeMailGroup(id);
          setLocalMailRows((current) => current.filter((row) => String(row.Id) !== id));
        }
        await refreshMailList();
      }}
      onRefresh={() => void refreshMailList()}
    />
  );
}

function MailListPage({ active, localMailRows, mailRows, onClearStatus, onCreate, onDelete, onEdit, onRefresh, recordTab, serverOptions, setRecordTab, setUserIdQuery, status, userIdQuery }: { active: MailSectionKey; localMailRows: Record<string, unknown>[]; mailRows: Record<string, unknown>[]; onClearStatus: () => void; onCreate: () => void; onDelete: (id: string) => Promise<void>; onEdit: (row: Record<string, unknown>) => void | Promise<void>; onRefresh: () => void; recordTab: "mail" | "claim"; serverOptions: ServerOption[]; setRecordTab: (tab: "mail" | "claim") => void; setUserIdQuery: (value: string) => void; status: string; userIdQuery: string }) {
  const global = active === "mailGlobal";
  const groupedMailIds = new Set(localMailRows.flatMap((row) => mailGroupIds(row)));
  const mergedRows = [...localMailRows, ...mailRows.filter((row) => !groupedMailIds.has(String(row.Id ?? row.id ?? "")))];
  const rows = mergedRows.filter((row) => {
    const targetIds = getArray(row.TargetID);
    const typ = Number(row.Typ);
    const isServerMail = typ === 2;
    const isPersonalMail = row.__mailListType === "personal" || typ === 3 || (!typ && targetIds.length > 0);
    if (global && isPersonalMail) return false;
    if (!global && !isPersonalMail) return false;
    if (!userIdQuery.trim()) return true;
    return formatCell(row.TargetID).includes(userIdQuery.trim());
  }).sort(compareMailRowsNewestFirst);
  const listColumns = ["ID", "邮件名称", "目标", "状态", "时间", "操作"];

  if (global && recordTab === "claim") {
    return (
      <section className="mail-page">
        <div className="mail-tabs"><button onClick={() => setRecordTab("mail")} type="button">全局邮件</button><button className="active" type="button">领取记录</button></div>
        <div className="mail-filter-line"><label>用户 ID：<input value={userIdQuery} onChange={(event) => setUserIdQuery(event.target.value)} /></label><button onClick={onRefresh} type="button"><Search size={14} />Search</button></div>
        <MailDataTable columns={["邮件ID", "邮件名称", "是否查看", "是否领取", "查看时间", "领取时间", "生效时间", "过期时间"]} rows={[]} />
      </section>
    );
  }

  return (
    <section className="mail-page">
      {global && <div className="mail-tabs"><button className="active" type="button">全局邮件</button><button onClick={() => setRecordTab("claim")} type="button">领取记录</button></div>}
      {!global && <div className="mail-filter-line"><label>用户 ID：<input value={userIdQuery} onChange={(event) => setUserIdQuery(event.target.value)} /></label><button onClick={onRefresh} type="button"><Search size={14} />Search</button></div>}
      <section className="mail-table-card">
        <button className="mail-primary-button" onClick={onCreate} type="button">新建</button>
        <PaginatedMailDataTable
          columns={listColumns}
          rows={rows.map((row) => {
            const id = String(row.Id ?? row.id ?? "");
            const targetIds = getArray(row.TargetID);
            const typ = Number(row.Typ);
            const isServerMail = typ === 2;
            const typeName = isServerMail ? "区服" : row.__mailListType === "personal" || typ === 3 || (!typ && targetIds.length > 0) ? "个人" : "全局";
            const title = formatCell(row.Titel ?? row.Title ?? "自定义邮件");
            const state = Number(row.St);
            const statusText = row.__scheduled
              ? row.__scheduledStatus === "failed" ? `定时失败：${formatCell(row.__scheduledError)}` : "待发送"
              : mailRowClaimed(row) ? "已领取" : state > 0 ? `状态 ${state}` : "未领取";
            const rawCreatedAt = row.CreateTime ?? row.CreatedAt ?? row.Ct ?? row.createdAt;
            const createdAtFallback = row.__local || row.__scheduled ? formatTimestampValue(rawCreatedAt) : "暂无数据";
            const regEnd = row.RegtEnd ?? row.Regt;
            return {
              ID: <span className="mail-id-cell"><span>{id}</span><em>{typeName}</em></span>,
              邮件名称: title,
              目标: <MailTargetSummary row={row} serverOptions={serverOptions} />,
              状态: statusText,
              时间: <span className="mail-time-cell"><span>注册开始: {formatMailListTime(row.RegtBegin, "2020-01-01 00:00")}</span><span>注册结束: {formatMailListTime(regEnd, defaultMailRegEndDisplay())}</span><span>生效时间: {formatMailListTime(row.St, "立即生效")}</span><span>过期时间: {formatMailListTime(row.Et, "2050-12-31 23:59")}</span><span>创建时间: {formatMailListTime(rawCreatedAt, createdAtFallback)}</span></span>,
              操作: <div className="mail-action-buttons"><button onClick={() => void onEdit(row)} type="button">编辑</button><button onClick={() => void onDelete(id)} type="button">撤回</button></div>,
            };
          })}
        />
      </section>
      <TransientToast kind={status.includes("失败") || status.includes("错误") ? "error" : "success"} message={status} onClose={onClearStatus} />
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

function PaginatedMailDataTable({ columns, pageSize = 20, rows }: { columns: string[]; pageSize?: number; rows: Array<Record<string, React.ReactNode>> }) {
  const [page, setPage] = React.useState(1);
  const pageCount = Math.max(1, Math.ceil(rows.length / pageSize));
  const safePage = Math.min(page, pageCount);
  React.useEffect(() => {
    setPage(1);
  }, [rows.length, pageSize]);
  React.useEffect(() => {
    if (page !== safePage) setPage(safePage);
  }, [page, safePage]);
  const start = rows.length ? (safePage - 1) * pageSize : 0;
  const pageRows = rows.slice(start, start + pageSize);
  const from = rows.length ? start + 1 : 0;
  const to = Math.min(start + pageSize, rows.length);
  return (
    <>
      <MailDataTable columns={columns} rows={pageRows} />
      <div className="table-pagination">
        <span>第 {from}-{to} 条/总共 {rows.length} 条</span>
        <button disabled={safePage <= 1} onClick={() => setPage((current) => Math.max(1, current - 1))} type="button">‹</button>
        <strong>{safePage}</strong>
        <button disabled={safePage >= pageCount} onClick={() => setPage((current) => Math.min(pageCount, current + 1))} type="button">›</button>
      </div>
    </>
  );
}

function MailTargetSummary({ row, serverOptions }: { row: Record<string, unknown>; serverOptions: ServerOption[] }) {
  const targetSummary = String(row.__targetSummary ?? "").trim();
  const targetIds = getArray(row.TargetID);
  const typ = Number(row.Typ);
  const isServerMail = typ === 2;
  const isPersonalMail = typ === 3 || (!typ && targetIds.length > 0);
  const targetText = targetIds.length
    ? targetSummary || (isServerMail
      ? targetIds.map((targetId) => serverOptions.find((server) => server.id === Number(targetId))?.name ?? gameServerDisplayName(targetId)).join(", ")
      : `用户 ${formatCell(row.TargetID)}`)
    : isPersonalMail
      ? "未填写用户"
      : "全服";
  const platforms = getArray(row.Platform).map((item) => Number(item)).filter((item) => item === 1 || item === 2);
  const versions = formatVersionConditionRows(getArray(row.__conditionRows));
  const fallbackVersions = versions.length ? versions : formatVersionConditionList(getArray(row.Version)).filter(Boolean);
  const regBegin = Number(row.RegtBegin);
  const regEnd = Number(row.RegtEnd ?? row.Regt);
  const tags = [
    { label: "目标", value: targetText },
    platforms.length ? { label: "系统", value: platforms.map((item) => item === 1 ? "GooglePlay" : "iOS").join(", ") } : null,
    fallbackVersions.length ? { label: "版本", value: fallbackVersions.join(", ") } : null,
    regBegin > 0 ? { label: "注册开始", value: formatTimestamp(regBegin) } : null,
    regEnd > 0 ? { label: "注册结束", value: formatTimestamp(regEnd) } : null,
  ].filter((item): item is { label: string; value: string } => Boolean(item));
  return <span className="mail-target-summary">{tags.map((item) => <span key={item.label}><em>{item.label}</em>{item.value}</span>)}</span>;
}

function mailRowCreateSeconds(row: Record<string, unknown>) {
  return parseTimeSeconds(row.CreateTime ?? row.CreatedAt ?? row.Ct ?? row.createdAt ?? row.CreateAt);
}

function mailRowIdBigInt(row: Record<string, unknown>) {
  const raw = String(row.Id ?? row.id ?? "").replace(/\D/g, "");
  if (!raw) return 0n;
  try {
    return BigInt(raw);
  } catch {
    return 0n;
  }
}

function compareMailRowsNewestFirst(a: Record<string, unknown>, b: Record<string, unknown>) {
  const timeDiff = mailRowCreateSeconds(b) - mailRowCreateSeconds(a);
  if (timeDiff !== 0) return timeDiff;
  const idA = mailRowIdBigInt(a);
  const idB = mailRowIdBigInt(b);
  if (idA === idB) return 0;
  return idB > idA ? 1 : -1;
}

function mailRowClaimed(row: Record<string, unknown>) {
  const type2 = Number(row.Type2 ?? row.type2);
  if (Number.isFinite(type2) && (type2 & 0x04) !== 0) return true;
  const truthyKeys = ["__claimed", "Claimed", "claimed", "IsClaim", "isClaim", "IsClaimed", "isClaimed", "IsGet", "isGet", "Got", "got", "Received", "received", "Drawed", "drawed"];
  if (truthyKeys.some((key) => {
    const value = row[key];
    return value === true || value === "true" || Number(value) === 1;
  })) return true;
  const timeKeys = ["ClaimTime", "claimTime", "GetTime", "getTime", "ReceiveTime", "receiveTime", "DrawTime", "drawTime", "GotTime", "gotTime"];
  if (timeKeys.some((key) => Number(row[key]) > 0 || (typeof row[key] === "string" && row[key] !== "" && row[key] !== "0" && row[key] !== "暂无数据"))) return true;
  const statusText = String(row.Status ?? row.status ?? row.State ?? row.state ?? row.ClaimStatus ?? row.claimStatus ?? "");
  if (/未领取|未领|unclaimed|not\s+(?:claimed|received)/i.test(statusText)) return false;
  return /已领取|领取|claimed|received|got/i.test(statusText);
}

function mailRowClaimStateKnown(row: Record<string, unknown>) {
  const keys = [
    "Type2", "type2", "__claimed", "Claimed", "claimed", "IsClaim", "isClaim", "IsClaimed", "isClaimed",
    "IsGet", "isGet", "Got", "got", "Received", "received", "Drawed", "drawed", "ClaimTime", "claimTime",
    "GetTime", "getTime", "ReceiveTime", "receiveTime", "DrawTime", "drawTime", "GotTime", "gotTime",
    "Status", "status", "State", "state", "ClaimStatus", "claimStatus",
  ];
  return keys.some((key) => Object.prototype.hasOwnProperty.call(row, key));
}

function formatServerIdList(value?: string) {
  return String(value ?? "")
    .split(/[\s,，;；]+/)
    .map((item) => item.trim())
    .filter(Boolean)
    .map(gameServerDisplayName)
    .join(", ");
}

function formatServerRange(ids: number[]) {
  const sorted = ids.filter((id) => Number.isFinite(id)).sort((a, b) => a - b);
  if (!sorted.length) return "";
  const first = gameServerDisplayName(sorted[0]);
  const last = gameServerDisplayName(sorted[sorted.length - 1]);
  return first === last ? first : `${first} - ${last}`;
}

function formatPlatformList(value?: string) {
  return String(value ?? "")
    .split(/[\s,，;；]+/)
    .map((item) => Number(item.trim()))
    .filter((item) => item === 1 || item === 2)
    .map((item) => item === 1 ? "GooglePlay" : "iOS")
    .join(", ");
}

function conditionOperatorOptions(field: string) {
  if (field === "version") return ["=", "<", "<="];
  if (field === "regTime") return [">=", "<="];
  if (field === "server") return ["=", ">=", "<="];
  return ["="];
}

function serverConditionIds(rows: Array<{ field: string; op: string; value: string }>, serverOptions: ServerOption[]) {
  const serverRows = rows.filter((row) => row.field === "server" && row.value.trim());
  if (!serverRows.length) return [];
  const allServerIds = serverOptions.map((server) => Number(server.id)).filter((id) => Number.isFinite(id)).sort((a, b) => a - b);
  if (!allServerIds.length && serverRows.every((row) => row.op === "=")) return Array.from(new Set(serverRows.flatMap((row) => toFlexibleNumberArray(row.value)))).sort((a, b) => a - b);
  let selected = new Set(allServerIds);
  for (const row of serverRows) {
    const values = toFlexibleNumberArray(row.value);
    if (!values.length) return [];
    let rowIds: number[];
    if (row.op === "=") {
      rowIds = values;
    } else {
      if (!allServerIds.length) return [];
      const boundary = row.op === ">=" ? Math.min(...values) : Math.max(...values);
      rowIds = allServerIds.filter((serverId) => row.op === ">=" ? serverId >= boundary : serverId <= boundary);
    }
    const rowSet = new Set(rowIds);
    selected = new Set([...selected].filter((serverId) => rowSet.has(serverId)));
  }
  return [...selected].sort((a, b) => a - b);
}

function MailEditor({ canUploadItemTable, global, initialMail, items, onBack, onSubmit, onUploadItemTable, rewardTemplates, serverOptions, templates }: { canUploadItemTable: boolean; global: boolean; initialMail?: Record<string, unknown>; items: ItemOption[]; onBack: () => void; onSubmit: (body: unknown) => Promise<string | void>; onUploadItemTable: (file: File) => Promise<void>; rewardTemplates: RewardTemplate[]; serverOptions: ServerOption[]; templates: MailTemplate[] }) {
  const now = new Date();
  const initialTyp = Number(initialMail?.Typ);
  const initialTargetIds = getArray(initialMail?.TargetID);
  const [mailType, setMailType] = React.useState(initialMail ? initialTyp === 3 || (!initialTyp && initialTargetIds.length > 0) ? "personal" : "global" : global ? "global" : "personal");
  const [templateId, setTemplateId] = React.useState("custom");
  const [rewardTemplateId, setRewardTemplateId] = React.useState("custom");
  const [rewardMode, setRewardMode] = React.useState<"none" | "reward">(() => getArray(initialMail?.ItemLst).length ? "reward" : "none");
  const [title, setTitle] = React.useState(String(initialMail?.Titel ?? initialMail?.Title ?? ""));
  const [body, setBody] = React.useState(String(initialMail?.Body ?? ""));
  const [targetIds, setTargetIds] = React.useState(initialTyp === 3 ? initialTargetIds.join(",") : "");
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
  const [filterRows, setFilterRows] = React.useState<Array<{ id: number; field: string; op: string; value: string }>>(() => {
    if (!initialMail) return [];
    const savedConditionRows = getArray(initialMail.__conditionRows)
      .map((row) => getObject(row))
      .filter((row): row is Record<string, unknown> => Boolean(row))
      .map((row, index) => ({
        id: index + 1,
        field: String(row.field ?? ""),
        op: String(row.op ?? "="),
        value: String(row.value ?? ""),
      }))
      .filter((row) => row.field && row.value.trim());
    if (savedConditionRows.length) return savedConditionRows;
    const rows: Array<{ id: number; field: string; op: string; value: string }> = [];
    let id = 1;
    if (initialTyp === 2 && initialTargetIds.length) rows.push({ id: id++, field: "server", op: "=", value: initialTargetIds.join(",") });
    const platforms = getArray(initialMail.Platform).map(String).filter(Boolean);
    if (platforms.length) rows.push({ id: id++, field: "system", op: "=", value: platforms.join(",") });
    const versions = formatVersionConditionList(getArray(initialMail.Version)).filter(Boolean);
    if (versions.length) rows.push({ id: id++, field: "version", op: "=", value: versions.join(",") });
    const regBegin = Number(initialMail.RegtBegin);
    const regEnd = Number(initialMail.RegtEnd ?? initialMail.Regt);
    if (regBegin > 0) rows.push({ id: id++, field: "regTime", op: ">=", value: secondsToDatetimeLocal(regBegin) });
    if (regEnd > 0) rows.push({ id: id++, field: "regTime", op: "<=", value: secondsToDatetimeLocal(regEnd) });
    return rows;
  });
  const [error, setError] = React.useState("");
  const [submitting, setSubmitting] = React.useState(false);
  const isGlobalMail = mailType === "global";
  const defaultFilterValue = (field: string, op = "=") => {
    if (field === "regTime") return op === "<=" ? defaultRegEndTime() : MAIL_DEFAULT_REG_BEGIN;
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
    if (!body.trim()) {
      setError("邮件内容为空");
      return;
    }
    const endSeconds = parseDatetimeLocalSeconds(endTime);
    if (!endSeconds) {
      setError("请选择有效的过期时间");
      return;
    }
    const nowSeconds = Math.floor(Date.now() / 1000);
    const rewardValidation = rewardMode === "reward" ? validateRewardRows(rewards, items) : { ok: true, itemList: [] as number[] };
    if (!rewardValidation.ok) {
      setError(rewardValidation.message ?? "请填写有效的奖励道具和数量，或选择无奖励");
      return;
    }
    const conditionRows = filterRows.filter((row) => isGlobalMail || row.field !== "server");
    const versionList = toVersionConditionArray(conditionRows);
    const platformList = conditionRows.filter((row) => row.field === "system").flatMap((row) => toPlatformNumberArray(row.value));
    const serverTargetIds = isGlobalMail ? serverConditionIds(conditionRows, serverOptions) : [];
    const regBeginValues = conditionRows.filter((row) => row.field === "regTime" && row.op === ">=" && row.value).map((row) => parseDatetimeLocalSeconds(dateToDatetimeLocal(row.value)));
    const regEndValues = conditionRows.filter((row) => row.field === "regTime" && row.op === "<=" && row.value).map((row) => parseDatetimeLocalSeconds(dateToDatetimeLocal(row.value, true)));
    const hasRegCondition = conditionRows.some((row) => row.field === "regTime" && row.value.trim());
    const regBeginSeconds = regBeginValues.length ? Math.max(...regBeginValues) : 0;
    const regEndSeconds = regEndValues.length ? Math.min(...regEndValues) : hasRegCondition ? parseDatetimeLocalSeconds(defaultRegEndTime()) : 0;
    if (regBeginValues.some((value) => !value) || regEndValues.some((value) => !value)) {
      setError("请选择有效的注册时间区间");
      return;
    }
    if (regBeginSeconds && regEndSeconds && regEndSeconds <= regBeginSeconds) {
      setError("注册结束时间必须晚于注册开始时间");
      return;
    }
    if (conditionRows.some((row) => row.field === "system" && row.op !== "=" && row.value.trim())) {
      setError("系统条件只支持等于");
      return;
    }
    const emptyCondition = conditionRows.find((row) => ["system", "version", "server"].includes(row.field) && !row.value.trim());
    if (emptyCondition) {
      const labels: Record<string, string> = { system: "系统", version: "版本", server: "区服" };
      setError(`${labels[emptyCondition.field] ?? "条件"}未填写`);
      return;
    }
    if (conditionRows.some((row) => row.field === "version" && row.value.trim() && !toVersionConditionArray([row]).length)) {
      setError("APP版本请填写 x.x、x.x.x 或 x.x.x.x 格式，例如 1.8 或 1.8.0.0");
      return;
    }
    if (conditionRows.some((row) => row.field === "system" && row.value.trim() && !toPlatformNumberArray(row.value).length)) {
      setError("系统请选择 GooglePlay 或 iOS");
      return;
    }
    if (isGlobalMail && conditionRows.some((row) => row.field === "server" && row.op !== "=" && row.value.trim()) && !serverOptions.length) {
      setError("区服范围条件需要先获取到区服列表");
      return;
    }
    if (isGlobalMail && conditionRows.some((row) => row.field === "server" && row.value.trim()) && !serverTargetIds.length) {
      setError("没有符合条件的区服");
      return;
    }
    if (isGlobalMail && filterRows.some((row) => row.field === "server" && !row.value.trim())) {
      setError("区服未填写");
      return;
    }
    const allServerIds = serverOptions.map((server) => Number(server.id)).filter((serverId) => Number.isFinite(serverId)).sort((a, b) => a - b);
    const serverTargetsCoverAll = isGlobalMail && serverTargetIds.length > 0 && allServerIds.length > 0 && serverTargetIds.length === allServerIds.length && serverTargetIds.every((serverId, index) => serverId === allServerIds[index]);
    const startSeconds = sendMode === "scheduled" ? parseDatetimeLocalSeconds(startTime) : 0;
    if (sendMode === "scheduled" && !startSeconds) {
      setError("请选择有效的定时发送时间");
      return;
    }
    if (sendMode === "scheduled" && startSeconds <= nowSeconds + 5) {
      setError("定时发送时间必须晚于当前时间");
      return;
    }
    if (sendMode === "scheduled" && endSeconds <= startSeconds) {
      setError("过期时间必须晚于生效时间");
      return;
    }
    if (sendMode === "scheduled" && endSeconds <= startSeconds + 10 * 60) {
      setError("过期时间间隔太短，请将过期时间设置为发送时间至少 11 分钟后");
      return;
    }
    if (sendMode !== "scheduled" && endSeconds <= nowSeconds + 10 * 60) {
      setError("过期时间间隔太短，请将过期时间设置为发送时间至少 11 分钟后");
      return;
    }
    const mailTyp = isGlobalMail ? serverTargetIds.length && !serverTargetsCoverAll ? 2 : 1 : 3;
    const mailTargets = isGlobalMail ? serverTargetsCoverAll ? [] : serverTargetIds : targets;
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
        LangLst: selectedTemplateContents ? mailLangListPayload(selectedTemplateContents) : [],
        BodyData: [],
        BodyData2: [],
        ItemLst: rewardMode === "reward" ? rewardValidation.itemList : [],
        Platform: platformList,
        Version: versionList,
        __conditionRows: conditionRows,
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
                return (
                  <div className="mail-condition-row" key={row.id}>
                    <select value={row.field} onChange={(event) => {
                      const nextField = !isGlobalMail && event.target.value === "server" ? "version" : event.target.value;
                      const nextOp = nextField === "regTime" ? ">=" : "=";
                      setFilterRows((current) => current.map((item) => item.id === row.id ? { ...item, field: nextField, op: nextOp, value: defaultFilterValue(nextField, nextOp) } : item));
                    }}>
                      {filterFieldOptions.map((option) => <option disabled={option.disabled} key={option.value} value={option.value}>{option.label}</option>)}
                    </select>
                    {["regTime", "version", "server"].includes(row.field) ? (
                      <select className="mail-condition-expression" value={row.op} onChange={(event) => {
                        const nextOp = event.target.value;
                        setFilterRows((current) => current.map((item) => item.id === row.id ? { ...item, op: nextOp, value: item.value || defaultFilterValue(item.field, nextOp) } : item));
                      }}>
                        {conditionOperatorOptions(row.field).map((op) => <option key={op} value={op}>{op}</option>)}
                      </select>
                    ) : (
                      <select className="mail-condition-expression" value="=" onChange={() => undefined}>
                        <option value="=">=</option>
                      </select>
                    )}
                    {row.field === "system" ? (
                      <select value={row.value} onChange={(event) => setFilterRows((current) => current.map((item) => item.id === row.id ? { ...item, value: event.target.value } : item))}>
                        <option value="">请选择</option>
                        <option value="1">GooglePlay</option>
                        <option value="2">iOS</option>
                      </select>
                    ) : row.field === "regTime" ? (
                      <input type="datetime-local" value={dateToDatetimeLocal(row.value, row.op === "<=")} onChange={(event) => setFilterRows((current) => current.map((item) => item.id === row.id ? { ...item, value: event.target.value } : item))} />
                    ) : row.field === "server" && serverOptions.length ? (
                      <ServerCascadeSelect options={serverOptions} value={row.value} onChange={(nextValue) => setFilterRows((current) => current.map((item) => item.id === row.id ? { ...item, value: nextValue } : item))} />
                    ) : (
                      <input disabled={unsupported} value={row.value} onChange={(event) => setFilterRows((current) => current.map((item) => item.id === row.id ? { ...item, value: event.target.value } : item))} placeholder={unsupported ? "当前接口暂未开放" : row.field === "version" ? "例如 1.8 或 1.8.0.0" : row.field === "server" ? "例如 12 或 1,2" : "例如 1,2"} />
                    )}
                    <button className="mail-condition-remove" onClick={() => setFilterRows((current) => current.filter((item) => item.id !== row.id))} type="button">删除</button>
                  </div>
                );
              })}
              <button className="mail-add-condition" onClick={() => setFilterRows((current) => [...current, { id: Date.now() + current.length, field: "version", op: "=", value: "" }])} type="button">新增条件</button>
              <small className="mail-condition-hint">{isGlobalMail ? "多个条件同时填写时为且的关系。填写“游戏内区服”后，会按接口 Typ=2 将 TargetID 作为区服ID发送。" : "多个条件同时填写时为且的关系。个人邮件会先由GM后台筛选用户，再提交符合条件的用户ID。"}</small>
            </div>
          </div>
          <label className="mail-form-row"><span>邮件奖励类型</span><select value={rewardMode} onChange={(event) => setRewardMode(event.target.value === "reward" ? "reward" : "none")}><option value="none">无奖励</option><option value="reward">有奖励</option></select></label>
          {rewardMode === "reward" && <label className="mail-form-row"><span>奖励模板</span><select value={rewardTemplateId} onChange={(event) => setRewardTemplateId(event.target.value)}><option value="custom">自定义</option>{rewardTemplates.map((template) => <option key={template.id} value={template.id}>{template.title}</option>)}</select></label>}
          {rewardMode === "reward" && <RewardRows canUploadItemTable={canUploadItemTable} items={items} onUploadItemTable={onUploadItemTable} rewards={rewards} setRewards={setRewards} />}
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
  const [addCount, setAddCount] = React.useState("1");
  const [addCountTip, setAddCountTip] = React.useState("");
  const updateReward = (index: number, patch: Partial<MailRewardItem>) => {
    setRewards(rewards.map((reward, rewardIndex) => rewardIndex === index ? { ...reward, ...patch } : reward));
  };
  const deleteReward = (index: number) => {
    setRewards(rewards.length > 1 ? rewards.filter((_, rewardIndex) => rewardIndex !== index) : [{ itemId: "", count: "0" }]);
  };
  const normalizeAddCount = () => {
    const raw = Number.parseInt(addCount, 10);
    if (Number.isFinite(raw) && raw > 50) {
      setAddCount("50");
      setAddCountTip("单次批量增加上限为50");
      return 50;
    }
    const count = Math.max(1, raw || 1);
    setAddCount(String(count));
    setAddCountTip("");
    return count;
  };
  const addRewards = () => {
    const count = normalizeAddCount();
    setRewards([...rewards, ...Array.from({ length: count }, () => ({ itemId: "", count: "0" }))]);
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
      <div className="mail-form-row mail-add-reward-row"><span /><div className="mail-add-reward-control"><button className="mail-add-reward" onClick={addRewards} type="button">新增</button><input aria-label="新增奖励条目数量" inputMode="numeric" min={1} max={50} type="number" value={addCount} onBlur={normalizeAddCount} onChange={(event) => { setAddCount(event.target.value); setAddCountTip(""); }} /><small>条</small>{addCountTip && <em>{addCountTip}</em>}</div></div>
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
  return <section className="mail-page"><div className="mail-filter-line"><label>模板名称：<input value={query} onChange={(event) => setQuery(event.target.value)} /></label><button type="button"><Search size={14} />Search</button></div><section className="mail-table-card"><button className="mail-primary-button" onClick={onCreate} type="button">新建</button><MailDataTable columns={["ID", "名称", "创建时间", "更新时间", "操作"]} rows={rows.map((template) => ({ ID: template.id, 名称: <span className="mail-ellipsis-cell" title={template.name}>{template.name}</span>, 创建时间: formatTimestampValue(template.createdAt), 更新时间: formatTimestampValue(template.updatedAt), 操作: <div className="mail-action-buttons"><button onClick={() => onEdit(template)} type="button">编辑</button><button onClick={() => void onDelete(template)} type="button">删除</button></div> }))} /></section></section>;
}

function MailTemplateEditor({ endpoint = "/local-api/mail-templates", kind = "邮件", onBack, onSaved, template }: { endpoint?: string; kind?: "邮件" | "公告"; onBack: () => void; onSaved: (template?: MailTemplate) => void; template?: MailTemplate }) {
  const [name, setName] = React.useState(template?.name ?? "");
  const [error, setError] = React.useState("");
  const [saving, setSaving] = React.useState(false);
  const [activeLanguage, setActiveLanguage] = React.useState(defaultMailLanguage);
  const [contents, setContents] = React.useState<Record<string, MailTemplateContent>>(() => {
    return normalizeLanguageContents(template?.contents, template ? { title: template.title ?? "", body: template.body ?? "" } : undefined);
  });
  const activeContent = contents[activeLanguage] ?? { title: "", body: "" };
  const updateContent = (patch: Partial<MailTemplateContent>) => {
    setContents((current) => ({ ...current, [activeLanguage]: { ...(current[activeLanguage] ?? { title: "", body: "" }), ...patch } }));
  };
  const downloadTemplate = () => {
    const workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.aoa_to_sheet([
      ["", ...languageDefinitions.map((language) => language.label)],
      ["Title", ...languageDefinitions.map(() => "")],
      ["desc", ...languageDefinitions.map(() => "")],
    ]);
    XLSX.utils.book_append_sheet(workbook, sheet, `${kind}模板`);
    XLSX.writeFile(workbook, `${kind}模板导入模板.xlsx`);
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
    const englishContent = cleanedContents[defaultMailLanguage] ?? { title: "", body: "" };
    if (!englishContent.title || !englishContent.body) {
      setError(kind === "公告" ? "至少填写一个英文公告标题和英文公告内容" : "至少填写一个英语标题和英语内容");
      setActiveLanguage(defaultMailLanguage);
      return;
    }
    const filledContents = fillMissingLanguageContents(cleanedContents);
    setError("");
    setSaving(true);
    try {
      const primary = filledContents[defaultMailLanguage].title && filledContents[defaultMailLanguage].body ? filledContents[defaultMailLanguage] : Object.values(filledContents).find((content) => content.title && content.body) ?? { title: "", body: "" };
      const response = await fetch(endpoint, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ id: template?.id, name: cleanName, title: primary.title, body: primary.body, contents: filledContents }) });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const payload = (await response.json().catch(() => ({}))) as { template?: MailTemplate };
      const nowSeconds = Math.floor(Date.now() / 1000);
      onSaved(payload.template ?? { id: template?.id ?? `t-${Date.now()}`, name: cleanName, title: primary.title, body: primary.body, contents: filledContents, createdAt: template?.createdAt ?? String(nowSeconds), updatedAt: String(nowSeconds) });
    } catch (saveError) {
      setError(saveError instanceof Error ? `保存失败：${saveError.message}` : "保存失败");
    } finally {
      setSaving(false);
    }
  };
  return (
    <section className="mail-edit-page">
      <div className="mail-edit-card">
        <header>{kind}模板</header>
        <div className="mail-template-name"><label>模板名称<input maxLength={MAX_TEMPLATE_NAME_LENGTH} value={name} onChange={(event) => setName(event.target.value)} placeholder="请输入模板名称" /></label><small>{name.length}/{MAX_TEMPLATE_NAME_LENGTH}</small></div>
        {error && <div className="mail-template-alert">{error}</div>}
        <div className="mail-language-tabs">{languageDefinitions.map((language) => <button className={activeLanguage === language.label ? "active" : ""} key={language.id} onClick={() => setActiveLanguage(language.label)} type="button"><small>{language.id}</small>{language.label}</button>)}</div>
        <div className="mail-form mail-template-form">
          <label className="mail-form-row"><span>{kind}标题</span><input value={activeContent.title} onChange={(event) => updateContent({ title: event.target.value })} placeholder={`请输入${activeLanguage}${kind}标题`} /></label>
          <label className="mail-form-row mail-template-body"><span>{kind}内容</span><textarea value={activeContent.body} onChange={(event) => updateContent({ body: event.target.value })} placeholder={`请输入${activeLanguage}${kind}内容`} /></label>
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
          <div className="mail-form-actions"><button disabled={saving} onClick={() => void save()} type="button">{saving ? "保存中..." : "保存"}</button><button disabled={saving} onClick={onBack} type="button">取消</button></div>
        </div>
      </div>
    </section>
  );
}

function NoticeTemplatePage() {
  const [templates, setTemplates] = React.useState<MailTemplate[]>([]);
  const [query, setQuery] = React.useState("");
  const [view, setView] = React.useState<"list" | "edit">("list");
  const [editingTemplate, setEditingTemplate] = React.useState<MailTemplate | undefined>();
  const refresh = React.useCallback(async () => {
    const response = await fetch("/local-api/notice-templates");
    const payload = (await response.json().catch(() => ({}))) as { templates?: MailTemplate[] };
    setTemplates(payload.templates ?? []);
  }, []);
  React.useEffect(() => {
    void refresh();
  }, [refresh]);
  if (view === "edit") {
    return <MailTemplateEditor endpoint="/local-api/notice-templates" kind="公告" onBack={() => setView("list")} onSaved={() => { setView("list"); setEditingTemplate(undefined); void refresh(); }} template={editingTemplate} />;
  }
  return <MailTemplateList query={query} setQuery={setQuery} templates={templates} onCreate={() => { setEditingTemplate(undefined); setView("edit"); }} onEdit={(template) => { setEditingTemplate(template); setView("edit"); }} onDelete={async (template) => {
    if (!window.confirm(`确认删除公告模板「${template.name}」？`)) return;
    await fetch(`/local-api/notice-templates/${encodeURIComponent(template.id)}`, { method: "DELETE" });
    await refresh();
  }} />;
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
      const error = apiBusinessError(result);
      if (error) {
        setStatus(`礼包码保存失败：${error}`);
        throw new Error(`礼包码保存失败：${error}`);
      }
      setStatus("礼包码已保存");
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
        <PaginatedMailDataTable
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
      <TransientToast kind={status.includes("失败") || status.includes("错误") ? "error" : "success"} message={status} onClose={() => setStatus("")} />
    </section>
  );
}

function GiftEditor({ canUploadItemTable, items, onBack, onSubmit, onUploadItemTable }: { canUploadItemTable: boolean; items: ItemOption[]; onBack: () => void; onSubmit: (body: unknown) => Promise<void>; onUploadItemTable: (file: File) => Promise<void> }) {
  const now = new Date();
  const later = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);
  const [giftType, setGiftType] = React.useState("1");
  const [groupMode, setGroupMode] = React.useState<"oneToOne" | "oneToMany">("oneToMany");
  const [groupIds, setGroupIds] = React.useState("");
  const [codes, setCodes] = React.useState("");
  const [templateName, setTemplateName] = React.useState("");
  const [maxCount, setMaxCount] = React.useState("1");
  const [endTime, setEndTime] = React.useState(toDatetimeLocal(later));
  const [rewards, setRewards] = React.useState<MailRewardItem[]>([{ itemId: "", count: "0" }]);
  const [error, setError] = React.useState("");
  const [submitting, setSubmitting] = React.useState(false);

  const submit = async () => {
    if (submitting) return;
    const ids = codes.split(/[\n,，\s]+/).map((item) => item.trim()).filter(Boolean);
    const groups = groupIds.split(/[\n,，\s]+/).map((item) => Number(item.trim())).filter((item) => Number.isFinite(item));
    if (!groups.length) {
      setError("请输入组ID");
      return;
    }
    if (!ids.length) {
      setError("请输入兑换码");
      return;
    }
    if (groupMode === "oneToMany" && groups.length !== 1) {
      setError("一对多模式下，组ID只能填写一个");
      return;
    }
    if (groupMode === "oneToOne" && groups.length !== ids.length) {
      setError("一对一批量模式下，组ID数量必须和兑换码数量一致");
      return;
    }
    const count = Number(maxCount);
    if (!Number.isInteger(count) || count <= 0) {
      setError("请输入兑换码数量");
      return;
    }
    const rewardValidation = validateRewardRows(rewards, items);
    if (!rewardValidation.ok) {
      setError(rewardValidation.message ?? "请填写有效奖励");
      return;
    }
    if (!rewardValidation.itemList.length) {
      setError("请至少添加一个奖励");
      return;
    }
    const endSeconds = parseDatetimeLocalSeconds(endTime);
    if (!endSeconds) {
      setError("请选择到期时间");
      return;
    }
    setError("");
    setSubmitting(true);
    try {
      await onSubmit({
        Id: ids,
        Type: Number(giftType),
        Group: groupMode === "oneToOne" && groups.length > 1 ? groups : groups[0],
        Num: count,
        Et: endSeconds,
        ItemLst: rewardValidation.itemList,
        Desc: templateName.trim(),
      });
    } catch (submitError) {
      setError(submitError instanceof Error ? submitError.message : "礼包码保存失败");
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <section className="gift-edit-page">
      <div className="gift-edit-card">
        <label className="gift-form-row"><span>兑换码类型</span><select value={giftType} onChange={(event) => setGiftType(event.target.value)}><option value="1">1（每个角色可兑换一次）</option><option value="0">0（普通兑换码）</option></select></label>
        <div className="gift-form-row gift-mode-row"><span>组ID模式</span><RadioPill checked={groupMode === "oneToMany"} label="一对多" onChange={() => setGroupMode("oneToMany")} /><RadioPill checked={groupMode === "oneToOne"} label="一对一批量" onChange={() => setGroupMode("oneToOne")} /></div>
        <div className="gift-help">
          <strong>说明：</strong>
          <p>一对多：一个组ID对应多个兑换码，组ID只填一个，兑换码可填多个。</p>
          <p>一对一批量：一个组ID对应一个兑换码，组ID行数必须和兑换码行数一致。</p>
        </div>
        <label className="gift-form-row gift-textarea-row"><span>组ID</span><textarea value={groupIds} onChange={(event) => setGroupIds(event.target.value)} placeholder={groupMode === "oneToOne" ? "请输入组ID，每行一个，与兑换码顺序一一对应" : "请输入一个组ID"} /></label>
        <label className="gift-form-row gift-textarea-row"><span>兑换码</span><textarea value={codes} onChange={(event) => setCodes(event.target.value)} placeholder={groupMode === "oneToOne" ? "请输入兑换码，每行一个，与组ID顺序一一对应" : "请输入兑换码，每行一个"} /></label>
        <label className="gift-form-row"><span>数量</span><input inputMode="numeric" value={maxCount} onChange={(event) => setMaxCount(event.target.value)} placeholder="数量" /></label>
        <label className="gift-form-row"><span>备注</span><input value={templateName} onChange={(event) => setTemplateName(event.target.value)} placeholder="请输入备注" /></label>
        <RewardRows canUploadItemTable={canUploadItemTable} items={items} onUploadItemTable={onUploadItemTable} rewards={rewards} setRewards={setRewards} />
        <label className="gift-form-row"><span>过期时间</span><input type="datetime-local" value={endTime} onChange={(event) => setEndTime(event.target.value)} /><em>北京时间 {formatBeijingTime(endTime)}</em></label>
        {error && <div className="mail-form-error gift-form-error">{error}</div>}
        <div className="gift-form-actions"><button disabled={submitting} onClick={() => void submit()} type="button">{submitting ? "保存中..." : "确定"}</button><button disabled={submitting} onClick={onBack} type="button">取消</button></div>
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
  return <section className="mail-page"><section className="mail-table-card"><button className="mail-primary-button" onClick={onCreate} type="button">新建</button><MailDataTable columns={["title", "创建时间", "更新时间", "操作"]} rows={templates.map((template) => ({ title: <span className="mail-ellipsis-cell" title={template.title}>{template.title}</span>, 创建时间: formatTimestampValue(template.createdAt), 更新时间: formatTimestampValue(template.updatedAt), 操作: <div className="mail-action-buttons"><button onClick={() => onEdit(template)} type="button">编辑</button><button onClick={() => void onDelete(template)} type="button">删除</button></div> }))} /></section></section>;
}

function RewardTemplateEditor({ canUploadItemTable, items, onBack, onSaved, onUploadItemTable, template }: { canUploadItemTable: boolean; items: ItemOption[]; onBack: () => void; onSaved: (template: RewardTemplate) => void; onUploadItemTable: (file: File) => Promise<void>; template?: RewardTemplate }) {
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
    if (cleanTitle.length > MAX_TEMPLATE_NAME_LENGTH) {
      setError(`奖励模板标题最多${MAX_TEMPLATE_NAME_LENGTH}个字符`);
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
      const payload = (await response.json().catch(() => ({}))) as { template?: RewardTemplate; error?: string };
      if (!response.ok || !payload.template) throw new Error(payload.error || `HTTP ${response.status}`);
      onSaved(payload.template);
    } catch (saveError) {
      setError(saveError instanceof Error ? `保存失败：${saveError.message}` : "保存失败");
    } finally {
      setSaving(false);
    }
  };
  return <section className="mail-edit-page"><div className="mail-edit-card reward-template-card"><header>奖励模板</header><div className="mail-form"><label className="mail-form-row reward-template-title-row"><span>标题</span><input maxLength={MAX_TEMPLATE_NAME_LENGTH} value={title} onChange={(event) => setTitle(event.target.value)} /><small>{title.length}/{MAX_TEMPLATE_NAME_LENGTH}</small></label><RewardRows canUploadItemTable={canUploadItemTable} items={items} onUploadItemTable={onUploadItemTable} rewards={rewards} setRewards={setRewards} />{error && <div className="mail-form-error">{error}</div>}<div className="mail-form-actions"><button disabled={saving} onClick={() => void save()} type="button">{saving ? "保存中..." : "保存"}</button><button disabled={saving} onClick={onBack} type="button">取消</button></div></div></div></section>;
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

function formatMailListTime(value: unknown, fallback: string) {
  const formatted = typeof value === "string" && /^\d{4}-\d{2}-\d{2}/.test(value) ? formatTimestampValue(value) : formatTimestamp(value);
  return formatted === "暂无数据" ? fallback : formatted;
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

function TransientToast({ kind = "info", message, onClose }: { kind?: "info" | "success" | "error"; message: string; onClose: () => void }) {
  React.useEffect(() => {
    if (!message) return undefined;
    const timer = window.setTimeout(onClose, 2400);
    return () => window.clearTimeout(timer);
  }, [message, onClose]);
  if (!message) return null;
  return <div className={`transient-toast ${kind}`} role="status">{message}</div>;
}

function noticeConditionsFromFields(notice: Partial<NoticeConfig>): ConditionRow[] {
  const rows: ConditionRow[] = [];
  let id = 1;
  if (notice.platforms?.trim()) rows.push({ id: id++, field: "system", op: "=", value: notice.platforms });
  if (notice.versions?.trim()) rows.push({ id: id++, field: "version", op: "=", value: notice.versions });
  if (Number(notice.typ) === 1 && notice.sid?.trim()) rows.push({ id: id++, field: "server", op: "=", value: notice.sid });
  if (notice.regBegin?.trim()) rows.push({ id: id++, field: "regTime", op: ">=", value: notice.regBegin });
  if (notice.regEnd?.trim()) rows.push({ id: id++, field: "regTime", op: "<=", value: notice.regEnd });
  return rows;
}

function noticeDefaultConditionValue(field: string, op = "=") {
  if (field === "regTime") return op === "<=" ? NOTICE_DEFAULT_REG_END : NOTICE_DEFAULT_REG_BEGIN;
  return "";
}

function NoticePage({ postWithToken }: { postWithToken: (endpoint: string, body: unknown) => Promise<ApiPostResponse> }) {
  const [notices, setNotices] = React.useState<NoticeConfig[]>([]);
  const [serverOptions, setServerOptions] = React.useState<ServerOption[]>([]);
  const [noticeTemplates, setNoticeTemplates] = React.useState<MailTemplate[]>([]);
  const [editing, setEditing] = React.useState<NoticeConfig | null>(null);
  const [form, setForm] = React.useState<NoticeConfig>({ slot: 1, templateName: "", title: "", body: "", contents: emptyLanguageContents(), imagePath: "", typ: 0, sid: "", regBegin: "", regEnd: "", platforms: "", versions: "", conditions: [] });
  const [noticeDrafts, setNoticeDrafts] = React.useState<NoticeConfig[]>([]);
  const [saving, setSaving] = React.useState(false);
  const [status, setStatus] = React.useState("");
  const palettes = ["blue", "green", "orange"];
  const emptyNotice = (slot: number): NoticeConfig => ({ slot, templateName: "", title: "", body: "", contents: emptyLanguageContents(), imagePath: "", typ: 0, sid: "", regBegin: "", regEnd: "", platforms: "", versions: "", conditions: [] });
  const normalizeNotice = (notice: Partial<NoticeConfig>, slot = Number(notice.slot) || 1): NoticeConfig => ({
    slot,
    templateName: notice.templateName ?? "",
    title: notice.title ?? "",
    body: notice.body ?? "",
    contents: normalizeLanguageContents(notice.contents, { title: notice.title ?? "", body: notice.body ?? "" }),
    imagePath: notice.imagePath ?? "",
    typ: notice.typ ?? 0,
    sid: notice.sid ?? "",
    regBegin: notice.regBegin ?? "",
    regEnd: notice.regEnd ?? "",
    platforms: notice.platforms ?? "",
    versions: notice.versions ?? "",
    conditions: Array.isArray(notice.conditions) ? notice.conditions : noticeConditionsFromFields(notice),
  });

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

  React.useEffect(() => {
    void fetch("/local-api/game-servers")
      .then((response) => response.json())
      .then((payload: { servers?: ServerOption[] }) => setServerOptions(payload.servers ?? []))
      .catch(() => setServerOptions([]));
  }, []);

  React.useEffect(() => {
    void fetch("/local-api/notice-templates")
      .then((response) => response.json())
      .then((payload: { templates?: MailTemplate[] }) => setNoticeTemplates(payload.templates ?? []))
      .catch(() => setNoticeTemplates([]));
  }, []);

  React.useEffect(() => {
    if (!editing) return undefined;
    document.body.classList.add("modal-open");
    return () => {
      document.body.classList.remove("modal-open");
    };
  }, [editing]);

  const openEditor = (notice?: NoticeConfig) => {
    const next = notice ?? notices[0] ?? { slot: 1, title: "", body: "", imagePath: "", typ: 0, sid: "" };
    const drafts = [1, 2, 3].map((slot) => normalizeNotice(notices.find((item) => Number(item.slot) === slot) ?? emptyNotice(slot), slot));
    setNoticeDrafts(drafts);
    setEditing(next);
    setForm(normalizeNotice(next, Number(next.slot) || 1));
  };

  const switchNoticeSlot = (slot: number) => {
    const currentSlot = Number(form.slot) || 1;
    const nextDrafts = [1, 2, 3].map((itemSlot) => {
      if (itemSlot === currentSlot) return normalizeNotice(form, currentSlot);
      return normalizeNotice(noticeDrafts.find((notice) => Number(notice.slot) === itemSlot) ?? notices.find((notice) => Number(notice.slot) === itemSlot) ?? emptyNotice(itemSlot), itemSlot);
    });
    const nextForm = nextDrafts.find((notice) => Number(notice.slot) === slot) ?? emptyNotice(slot);
    setNoticeDrafts(nextDrafts);
    setForm(normalizeNotice(nextForm, slot));
  };

  const save = async () => {
    if (saving) return;
    const normalizedContents = fillMissingLanguageContents(form.contents, { title: form.title, body: form.body });
    const primaryNoticeContent = normalizedContents[defaultMailLanguage].title && normalizedContents[defaultMailLanguage].body ? normalizedContents[defaultMailLanguage] : Object.values(normalizedContents).find((content) => content.title && content.body) ?? { title: "", body: "" };
    const conditionRows = form.conditions ?? [];
    const versionList = toVersionConditionArray(conditionRows);
    const platformList = conditionRows.filter((row) => row.field === "system").flatMap((row) => toPlatformNumberArray(row.value));
    const sidValues = serverConditionIds(conditionRows, serverOptions);
    const regBeginValues = conditionRows.filter((row) => row.field === "regTime" && row.op === ">=" && row.value).map((row) => parseDatetimeLocalSeconds(dateToDatetimeLocal(row.value)));
    const regEndValues = conditionRows.filter((row) => row.field === "regTime" && row.op === "<=" && row.value).map((row) => parseDatetimeLocalSeconds(dateToDatetimeLocal(row.value, true)));
    const regBeginSeconds = regBeginValues.length ? Math.max(...regBeginValues) : parseDatetimeLocalSeconds(NOTICE_DEFAULT_REG_BEGIN);
    const regEndSeconds = regEndValues.length ? Math.min(...regEndValues) : parseDatetimeLocalSeconds(NOTICE_DEFAULT_REG_END);
    if (conditionRows.some((row) => row.field === "system" && row.op !== "=" && row.value.trim())) {
      setStatus("公告保存失败：系统条件只支持等于");
      return;
    }
    const emptyCondition = conditionRows.find((row) => ["system", "version", "server"].includes(row.field) && !row.value.trim());
    if (emptyCondition) {
      const labels: Record<string, string> = { system: "系统", version: "版本", server: "区服" };
      setStatus(`公告保存失败：${labels[emptyCondition.field] ?? "条件"}未填写`);
      return;
    }
    if (conditionRows.some((row) => row.field === "version" && row.value.trim() && !toVersionConditionArray([row]).length)) {
      setStatus("公告保存失败：版本请填写 x.x、x.x.x 或 x.x.x.x，例如 1.8 或 1.8.0.0");
      return;
    }
    if (conditionRows.some((row) => row.field === "system" && row.value.trim() && !toPlatformNumberArray(row.value).length)) {
      setStatus("公告保存失败：平台请选择 GooglePlay 或 iOS");
      return;
    }
    if (conditionRows.some((row) => row.field === "server" && row.op !== "=" && row.value.trim()) && !serverOptions.length) {
      setStatus("公告保存失败：区服范围条件需要先获取到区服列表");
      return;
    }
    if (conditionRows.some((row) => row.field === "server" && row.value.trim()) && !sidValues.length) {
      setStatus("公告保存失败：没有符合条件的区服");
      return;
    }
    if (regBeginValues.some((value) => !value) || regEndValues.some((value) => !value)) {
      setStatus("公告保存失败：请选择有效的注册时间区间");
      return;
    }
    if (regBeginSeconds && regEndSeconds && regEndSeconds <= regBeginSeconds) {
      setStatus("公告保存失败：注册结束时间必须晚于注册开始时间");
      return;
    }
    const effectiveForm = {
      ...form,
      imagePath: form.imagePath.trim() || NOTICE_DEFAULT_IMAGE,
      typ: sidValues.length ? 1 : 0,
      sid: sidValues.join(","),
      regBegin: secondsToDatetimeLocal(regBeginSeconds),
      regEnd: secondsToDatetimeLocal(regEndSeconds),
      platforms: platformList.length ? platformList.join(",") : "1,2",
      versions: versionList.length ? versionList.join(",") : NOTICE_DEFAULT_VERSION_RANGE,
      conditions: conditionRows,
    };
    const normalizedForm = normalizeNotice({ ...effectiveForm, title: primaryNoticeContent.title, body: primaryNoticeContent.body, contents: normalizedContents }, Number(effectiveForm.slot) || 1);
    setSaving(true);
    try {
      const merged = [1, 2, 3].map((slot) => normalizeNotice(slot === normalizedForm.slot ? normalizedForm : noticeDrafts.find((notice) => Number(notice.slot) === slot) ?? notices.find((notice) => Number(notice.slot) === slot) ?? emptyNotice(slot), slot));
      const result = await postWithToken("/gmNoticeAdd", configsToNoticePayload(merged));
      const error = apiBusinessError(result);
      if (error) {
        setStatus(`公告保存失败：${error}`);
        return;
      }
      setStatus("公告已保存到服务器");
      setEditing(null);
      await refresh();
    } catch (saveError) {
      setStatus(saveError instanceof Error ? `公告保存失败：${saveError.message}` : "公告保存失败");
    } finally {
      setSaving(false);
    }
  };

  const applyNoticeTemplate = (templateId: string) => {
    const template = noticeTemplates.find((item) => item.id === templateId);
    if (!template) {
      setForm({ ...form, templateName: "", title: "", body: "", contents: emptyLanguageContents() });
      return;
    }
    const contents = normalizeLanguageContents(template.contents, { title: template.title, body: template.body });
    const primary = contents[defaultMailLanguage] ?? Object.values(contents).find((content) => content.title || content.body) ?? { title: "", body: "" };
    setForm({ ...form, templateName: template.name, title: primary.title, body: primary.body, contents });
  };

  const deleteNotice = async (slot: number) => {
    if (!window.confirm(`确认删除公告 ${slot}？`)) return;
    const merged = [1, 2, 3].map((itemSlot) => normalizeNotice(itemSlot === slot ? emptyNotice(itemSlot) : notices.find((notice) => Number(notice.slot) === itemSlot) ?? emptyNotice(itemSlot), itemSlot));
    const result = await postWithToken("/gmNoticeAdd", configsToNoticePayload(merged));
    const error = apiBusinessError(result);
    if (error) {
      setStatus(`公告删除失败：${error}`);
      return;
    }
    setStatus(`公告 ${slot} 已删除`);
    if (Number(editing?.slot) === slot) setEditing(null);
    await refresh();
  };
  const noticeFilterFieldOptions = [
    { value: "system", label: "系统" },
    { value: "version", label: "app版本" },
    { value: "regTime", label: "注册时间" },
    { value: "server", label: "游戏内区服" },
  ];
  const noticeConditionRows = form.conditions ?? [];

  return (
    <section className="notice-page">
      <h2>公告</h2>
      <button className="notice-edit-button" onClick={() => openEditor()} type="button"><Bell size={15} />添加/修改公告</button>
      <div className="notice-grid">
        {[1, 2, 3].map((slot, index) => {
          const notice = notices.find((item) => Number(item.slot) === slot) ?? { slot, title: "", body: "", imagePath: "" };
          return (
            <article className={`notice-card ${palettes[index]}`} key={slot}>
              <button className="notice-delete-button" onClick={(event) => { event.stopPropagation(); void deleteNotice(slot); }} title={`删除公告 ${slot}`} type="button"><Trash2 size={14} />删除</button>
              <button className="notice-card-content" onClick={() => openEditor(notice)} type="button">
                <span className="notice-watermark">admin</span>
                <strong>公告 {slot}</strong>
                <h3>{notice.templateName || notice.title || "未配置公告模板"}</h3>
                <div className="notice-body-box">{notice.title || "暂无公告标题"}</div>
                <label>配图路径</label>
                <div className="notice-image-path">{notice.imagePath || "未配置配图路径"}</div>
                <div className="tag-row">
                  <small>{Number(notice.typ) === 1 ? `服务器：${formatServerIdList(notice.sid) || "未填写"}` : "范围：全部服务器"}</small>
                  <small>注册开始：{formatBeijingTime(notice.regBegin || NOTICE_DEFAULT_REG_BEGIN)}</small>
                  <small>注册结束：{formatBeijingTime(notice.regEnd || NOTICE_DEFAULT_REG_END)}</small>
                  <small>平台：{notice.platforms ? formatPlatformList(notice.platforms) : "全部平台"}</small>
                  <small>版本：{notice.versions || "全部版本"}</small>
                </div>
              </button>
            </article>
          );
        })}
      </div>
      <TransientToast kind={status.includes("失败") || status.includes("错误") ? "error" : "success"} message={status} onClose={() => setStatus("")} />
      {editing && (
        <div className="modal-backdrop" role="presentation">
          <section className="notice-modal" role="dialog" aria-modal="true">
            <header><strong>添加/修改公告</strong><button onClick={() => setEditing(null)} type="button">x</button></header>
            <div className="notice-form">
              <label>公告位置<select value={form.slot} onChange={(event) => switchNoticeSlot(Number(event.target.value))}><option value={1}>公告 1</option><option value={2}>公告 2</option><option value={3}>公告 3</option></select></label>
              <label>公告模板<select value={noticeTemplates.find((template) => template.name === form.templateName)?.id ?? ""} onChange={(event) => applyNoticeTemplate(event.target.value)}><option value="">请选择公告模板</option>{noticeTemplates.map((template) => <option key={template.id} value={template.id}>{template.name}</option>)}</select></label>
              <label>配图路径<input value={form.imagePath} onChange={(event) => setForm({ ...form, imagePath: event.target.value })} placeholder={`默认：${NOTICE_DEFAULT_IMAGE}`} /></label>
              <div className="notice-condition-block">
                <span>条件</span>
                <div className="mail-condition-list">
                  {noticeConditionRows.length === 0 && <div className="mail-condition-empty">默认无条件，公告会展示给全部目标。</div>}
                  {noticeConditionRows.map((row) => (
                    <div className="mail-condition-row" key={row.id}>
                      <select value={row.field} onChange={(event) => {
                        const nextField = event.target.value;
                        const nextOp = nextField === "regTime" ? ">=" : "=";
                        setForm({ ...form, conditions: noticeConditionRows.map((item) => item.id === row.id ? { ...item, field: nextField, op: nextOp, value: noticeDefaultConditionValue(nextField, nextOp) } : item) });
                      }}>
                        {noticeFilterFieldOptions.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}
                      </select>
                      {["regTime", "version", "server"].includes(row.field) ? (
                        <select className="mail-condition-expression" value={row.op} onChange={(event) => {
                          const nextOp = event.target.value;
                          setForm({ ...form, conditions: noticeConditionRows.map((item) => item.id === row.id ? { ...item, op: nextOp, value: item.value || noticeDefaultConditionValue(item.field, nextOp) } : item) });
                        }}>
                          {conditionOperatorOptions(row.field).map((op) => <option key={op} value={op}>{op}</option>)}
                        </select>
                      ) : (
                        <select className="mail-condition-expression" value="=" onChange={() => undefined}><option value="=">=</option></select>
                      )}
                      {row.field === "system" ? (
                        <select value={row.value} onChange={(event) => setForm({ ...form, conditions: noticeConditionRows.map((item) => item.id === row.id ? { ...item, value: event.target.value } : item) })}>
                          <option value="">请选择</option>
                          <option value="1">GooglePlay</option>
                          <option value="2">iOS</option>
                        </select>
                      ) : row.field === "regTime" ? (
                        <input type="datetime-local" value={dateToDatetimeLocal(row.value, row.op === "<=")} onChange={(event) => setForm({ ...form, conditions: noticeConditionRows.map((item) => item.id === row.id ? { ...item, value: event.target.value } : item) })} />
                      ) : row.field === "server" && serverOptions.length ? (
                        <ServerCascadeSelect options={serverOptions} value={row.value} onChange={(nextValue) => setForm({ ...form, conditions: noticeConditionRows.map((item) => item.id === row.id ? { ...item, value: nextValue } : item) })} />
                      ) : (
                        <input value={row.value} onChange={(event) => setForm({ ...form, conditions: noticeConditionRows.map((item) => item.id === row.id ? { ...item, value: event.target.value } : item) })} placeholder={row.field === "version" ? "例如 1.8 或 1.8.0.0" : row.field === "server" ? "例如 12 或 1,2" : "例如 1,2"} />
                      )}
                      <button className="mail-condition-remove" onClick={() => setForm({ ...form, conditions: noticeConditionRows.filter((item) => item.id !== row.id) })} type="button">删除</button>
                    </div>
                  ))}
                  <button className="mail-add-condition" onClick={() => setForm({ ...form, conditions: [...noticeConditionRows, { id: Date.now() + noticeConditionRows.length, field: "system", op: "=", value: "" }] })} type="button">新增条件</button>
                  <small className="mail-condition-hint">多个条件同时填写时为且的关系。游戏内区服、系统、版本和注册时间会转换成公告接口筛选条件。</small>
                </div>
              </div>
              {status && <div className="mail-form-error">{status}</div>}
              <footer><button disabled={saving} onClick={() => void save()} type="button">{saving ? "保存中..." : "保存"}</button><button disabled={saving} onClick={() => setEditing(null)} type="button">取消</button></footer>
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
    const sid = data?.[`Sid${slot}`];
    const title = String(data?.[`Titel${suffix}`] ?? "");
    const body = String(data?.[`Body${suffix}`] ?? "");
    const contents = parseNoticeLanguageContents(data, slot, { title, body });
    const platformValues = Array.isArray(platform) ? platform.map(Number).filter(Number.isFinite) : [];
    const usesAllPlatforms = platformValues.length === 2 && platformValues.includes(1) && platformValues.includes(2);
    const versionValues = Array.isArray(version) ? version.map(Number).filter(Number.isFinite) : [];
    const usesAllVersions = versionValues.length === 2 && versionValues.includes(0) && versionValues.includes(4294967295);
    const regBeginValue = Number(data?.[`RegtBegin${slot}`] ?? 0);
    const regEndValue = Number(data?.[`RegtEnd${slot}`] ?? 0);
    const defaultRegBeginSeconds = parseDatetimeLocalSeconds(NOTICE_DEFAULT_REG_BEGIN);
    const defaultRegEndSeconds = parseDatetimeLocalSeconds(NOTICE_DEFAULT_REG_END);
    const versionText = versionValues.length && !usesAllVersions ? formatVersionConditionList(versionValues).join(",") : "";
    const config = {
      slot,
      templateName: String(data?.[`TemplateName${slot}`] ?? data?.[`NoticeTemplate${slot}`] ?? ""),
      title,
      body,
      contents,
      imagePath: String(data?.[`Rs${slot}`] ?? ""),
      typ: Number(data?.[`Typ${slot}`] ?? 0),
      sid: Array.isArray(sid) ? sid.join(",") : "",
      regBegin: regBeginValue > 0 && regBeginValue !== defaultRegBeginSeconds ? secondsToDatetimeLocal(regBeginValue) : "",
      regEnd: regEndValue > 0 && regEndValue !== defaultRegEndSeconds ? secondsToDatetimeLocal(regEndValue) : "",
      platforms: platformValues.length && !usesAllPlatforms ? platformValues.join(",") : "",
      versions: versionText,
    };
    return {
      ...config,
      conditions: noticeConditionsFromFields(config),
    };
  });
}

function parseNoticeLanguageContents(data: Record<string, unknown> | null, slot: number, fallback: MailTemplateContent) {
  const normalized = normalizeLanguageContents(undefined, fallback);
  const rawLanguageData = data?.[`LangLst${slot}`] ?? data?.[`Lang${slot}`] ?? data?.[`Languages${slot}`] ?? data?.[`Language${slot}`];
  const titleList = getArray(data?.[`TitelLst${slot}`] ?? data?.[`TitleLst${slot}`]);
  const bodyList = getArray(data?.[`BodyLst${slot}`]);
  const applyRows = (rows: unknown[]) => {
    for (const row of rows) {
      const item = getObject(row);
      if (!item) continue;
      const language = languageDefinitions.find((candidate) => Number(item.id ?? item.Id ?? item.Language ?? item.LanguageId ?? item.languageId) === candidate.id) ?? matchLanguage(item.language ?? item.Language ?? item.code ?? item.Code);
      if (!language) continue;
      normalized[language.label] = { title: String(item.title ?? item.Title ?? item.Titel ?? ""), body: String(item.body ?? item.Body ?? "") };
    }
  };
  if (Array.isArray(rawLanguageData)) applyRows(rawLanguageData);
  if (titleList.length || bodyList.length) {
    languageDefinitions.forEach((language, index) => {
      normalized[language.label] = { title: String(titleList[index] ?? normalized[language.label].title ?? ""), body: String(bodyList[index] ?? normalized[language.label].body ?? "") };
    });
  }
  return normalized;
}

function configsToNoticePayload(configs: NoticeConfig[]) {
  const payload: Record<string, unknown> = {};
  for (const config of configs) {
    const slot = Number(config.slot);
    const suffix = slot === 1 ? "" : String(slot);
    const rawContents = normalizeLanguageContents(config.contents, { title: config.title, body: config.body });
    const hasContent = Boolean(
      String(config.title ?? "").trim()
      || String(config.body ?? "").trim()
      || Object.values(rawContents).some((content) => content.title.trim() || content.body.trim())
    );
    if (!hasContent) {
      payload[`Titel${suffix}`] = "";
      payload[`Body${suffix}`] = "";
      payload[`LangLst${slot}`] = [];
      payload[`Rs${slot}`] = "";
      payload[`Typ${slot}`] = 0;
      payload[`RegtBegin${slot}`] = 0;
      payload[`RegtEnd${slot}`] = 0;
      payload[`Sid${slot}`] = [];
      payload[`Platform${slot}`] = [];
      payload[`Version${slot}`] = [];
      continue;
    }
    const effectiveConfig = {
      ...config,
      imagePath: config.imagePath?.trim() || NOTICE_DEFAULT_IMAGE,
      regBegin: config.regBegin || NOTICE_DEFAULT_REG_BEGIN,
      regEnd: config.regEnd || NOTICE_DEFAULT_REG_END,
      platforms: config.platforms || "1,2",
      versions: config.versions || NOTICE_DEFAULT_VERSION_RANGE,
    };
    const contents = fillMissingLanguageContents(effectiveConfig.contents, { title: effectiveConfig.title, body: effectiveConfig.body });
    const primary = contents[defaultMailLanguage].title && contents[defaultMailLanguage].body ? contents[defaultMailLanguage] : Object.values(contents).find((content) => content.title || content.body) ?? { title: "", body: "" };
    const langList = mailLangListPayload(contents);
    payload[`Titel${suffix}`] = primary.title ?? "";
    payload[`Body${suffix}`] = primary.body ?? "";
    payload[`LangLst${slot}`] = langList;
    payload[`Rs${slot}`] = effectiveConfig.imagePath;
    const sid = toFlexibleNumberArray(effectiveConfig.sid);
    const isSpecifiedServer = Number(effectiveConfig.typ) === 1 && sid.length > 0;
    payload[`Typ${slot}`] = isSpecifiedServer ? 1 : 0;
    payload[`RegtBegin${slot}`] = effectiveConfig.regBegin ? parseDatetimeLocalSeconds(effectiveConfig.regBegin) : 0;
    payload[`RegtEnd${slot}`] = effectiveConfig.regEnd ? parseDatetimeLocalSeconds(effectiveConfig.regEnd) : 0;
    payload[`Sid${slot}`] = isSpecifiedServer ? sid : [];
    const platforms = toPlatformNumberArray(effectiveConfig.platforms);
    const versions = toNoticeVersionArray(effectiveConfig.versions);
    payload[`Platform${slot}`] = platforms.length ? platforms : [];
    payload[`Version${slot}`] = versions.length ? versions : [];
  }
  return payload;
}

function AccountPanel({ accounts, canManageAdmins, games, session, onAdd, onDelete, onUpdate, onClose }: { accounts: ManagedAccount[]; canManageAdmins: boolean; games: GameConfig[]; session: Session | null; onAdd: (account: ManagedAccount) => Promise<void>; onDelete: (accountId: number) => Promise<void>; onUpdate: (account: ManagedAccount) => Promise<void>; onClose: () => void }) {
  const fallbackGameKey = games[0] ? `${games[0].name}/${games[0].serverName}` : "";
  const currentGameKey = session ? `${session.game}/${session.serverName}` : fallbackGameKey;
  const emptyForm = { account: "", password: "", displayName: "", role: "运营", gameKeys: currentGameKey ? [currentGameKey] : [], permissions: ["用户查询", "日志审计"], isManager: false };
  const [form, setForm] = React.useState(emptyForm);
  const [editingAccount, setEditingAccount] = React.useState<ManagedAccount | null>(null);
  const [formError, setFormError] = React.useState("");
  const [formSuccess, setFormSuccess] = React.useState("");
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
    setFormSuccess("");
  };
  const resetForm = () => {
    setEditingAccount(null);
    setForm(emptyForm);
    setFormError("");
    setFormSuccess("");
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
              setFormSuccess("");
              const payload = { id: editingAccount?.id ?? Date.now(), account: form.account, password: form.password, displayName: form.displayName || form.account, role: form.role, games: form.isManager ? gameOptions.map((game) => game.value) : form.gameKeys, permissions: form.permissions, isManager: canManageAdmins && form.isManager, status: editingAccount?.status ?? "启用" };
              if (editingAccount) {
                await onUpdate(payload);
                resetForm();
                setFormSuccess("账号修改成功");
              } else {
                await onAdd(payload);
                setForm((current) => ({ ...current, account: "", password: "", displayName: "", gameKeys: [currentGameKey], isManager: false }));
                setFormSuccess("账号创建成功");
              }
            } catch (error) {
              setFormSuccess("");
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
            <TransientToast kind="success" message={formSuccess} onClose={() => setFormSuccess("")} />
            <TransientToast kind="error" message={formError} onClose={() => setFormError("")} />
          </form>
          <div className="managed-list account-managed-list">{accounts.length === 0 ? <div className="empty-table"><strong>暂无子账号</strong><span>请在左侧创建账号给其他人使用。</span></div> : accounts.map((account) => (
            <article className="managed-account account-managed-item" key={account.id}>
              <div><strong>{account.displayName}</strong><span>{account.account} / {account.role}{account.isManager ? " / 后台管理员" : ""}</span></div>
              <div className="tag-row">{account.isManager ? <small>全部游戏</small> : account.games.map((item) => <small key={item}>{item}</small>)}{account.permissions.slice(0, 3).map((item) => <small key={item}>{item}</small>)}</div>
              <div className="managed-actions"><button onClick={() => beginEdit(account)} type="button"><UserCog size={14} />编辑</button>{canManageAdmins ? <button onClick={async () => {
                try {
                  setFormError("");
                  setFormSuccess("");
                  await onDelete(account.id);
                  setFormSuccess("账号删除成功");
                } catch (error) {
                  setFormSuccess("");
                  setFormError(error instanceof Error ? error.message : "删除账号失败");
                }
              }} type="button"><Trash2 size={14} />删除</button> : <button disabled type="button"><Trash2 size={14} />仅admin可删</button>}</div>
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

function ServerCascadeSelect({ onChange, options, value }: { onChange: (value: string) => void; options: ServerOption[]; value: string }) {
  const [open, setOpen] = React.useState(false);
  const [openUp, setOpenUp] = React.useState(false);
  const containerRef = React.useRef<HTMLDivElement | null>(null);
  const groups = React.useMemo(() => {
    const normalized = options
      .map((server) => ({ id: Number(server.id), name: server.name || gameServerDisplayName(server.id) }))
      .filter((server) => Number.isFinite(server.id) && server.id > 0)
      .sort((a, b) => a.id - b.id);
    const glServers = normalized.filter((server) => server.id <= 123);
    const tkServers = normalized.filter((server) => server.id > 123);
    const result: Array<{ key: string; label: string; servers: Array<{ id: number; name: string }> }> = [];
    if (glServers.length) result.push({ key: "gl", label: "GL-1 - GL-123", servers: glServers });
    const maxTkNumber = tkServers.reduce((max, server) => Math.max(max, server.id - 123), 0);
    for (let start = 1; start <= maxTkNumber; start += 100) {
      const end = Math.min(start + 99, maxTkNumber);
      const servers = tkServers.filter((server) => {
        const tkNumber = server.id - 123;
        return tkNumber >= start && tkNumber <= end;
      });
      if (servers.length) result.push({ key: `tk-${start}`, label: `TK-${start} - TK-${end}`, servers });
    }
    return result;
  }, [options]);
  const [activeKey, setActiveKey] = React.useState(groups[0]?.key ?? "");
  React.useEffect(() => {
    if (!groups.some((group) => group.key === activeKey)) setActiveKey(groups[0]?.key ?? "");
  }, [activeKey, groups]);
  const activeGroup = groups.find((group) => group.key === activeKey) ?? groups[0];
  const selectedName = value
    ? value.includes(",") || value.includes("，") ? formatServerIdList(value) : options.find((server) => String(server.id) === String(value))?.name ?? gameServerDisplayName(value)
    : "";
  React.useEffect(() => {
    if (!open) return undefined;
    const closeOnOutsidePointerDown = (event: PointerEvent) => {
      if (containerRef.current?.contains(event.target as Node)) return;
      setOpen(false);
    };
    document.addEventListener("pointerdown", closeOnOutsidePointerDown);
    return () => document.removeEventListener("pointerdown", closeOnOutsidePointerDown);
  }, [open]);

  const toggleMenu = () => {
    if (open) {
      setOpen(false);
      return;
    }
    const triggerRect = containerRef.current?.getBoundingClientRect();
    const scrollContainer = containerRef.current?.closest(".notice-form");
    const boundaryRect = scrollContainer?.getBoundingClientRect();
    if (triggerRect) {
      const boundaryTop = boundaryRect?.top ?? 0;
      const boundaryBottom = boundaryRect?.bottom ?? window.innerHeight;
      const spaceAbove = triggerRect.top - boundaryTop;
      const spaceBelow = boundaryBottom - triggerRect.bottom;
      setOpenUp(spaceBelow < 324 && spaceAbove > spaceBelow);
    }
    setOpen(true);
  };

  return (
    <div className="server-cascade-select" ref={containerRef}>
      <button className="server-cascade-trigger" onClick={toggleMenu} type="button">
        <span>{selectedName || "请选择游戏内区服"}</span>
        <ChevronDown size={15} />
      </button>
      {open && (
        <div className={`server-cascade-menu${openUp ? " open-up" : ""}`}>
          <div className="server-cascade-groups">
            {groups.map((group) => (
              <button className={group.key === activeGroup?.key ? "active" : ""} key={group.key} onMouseEnter={() => setActiveKey(group.key)} type="button">{group.label}</button>
            ))}
          </div>
          <div className="server-cascade-options">
            {activeGroup?.servers.map((server) => (
              <button className={String(value) === String(server.id) ? "active" : ""} key={server.id} onClick={() => { onChange(String(server.id)); setOpen(false); }} type="button">{server.name}</button>
            ))}
          </div>
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
  const [formSuccess, setFormSuccess] = React.useState("");
  const loadImage = (file: File | undefined, key: "iconUrl" | "backgroundUrl") => {
    if (!file) return;
    if (!file.type.startsWith("image/")) {
      setFormError("请上传图片文件");
      setFormSuccess("");
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
    setFormSuccess("");
  };
  const beginEdit = (game: GameConfig) => {
    setEditingId(game.id ?? null);
    setForm({ ...emptyForm, ...game, iconUrl: game.iconUrl ?? "", backgroundUrl: game.backgroundUrl ?? "" });
    setFormError("");
    setFormSuccess("");
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
              setFormSuccess("");
              if (!form.serverAccount?.trim() || !form.serverPassword?.trim()) {
                setFormError("请填写该区服的服务端账号和密码");
                return;
              }
              if (editingId) {
                await onUpdate({ ...form, id: editingId });
                resetForm();
                setFormSuccess("区服修改成功");
              } else {
                await onAdd(form);
                setForm((current) => ({ ...emptyForm, name: current.name }));
                setFormSuccess("区服新增成功");
              }
            } catch (error) {
              setFormSuccess("");
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
            <TransientToast kind="success" message={formSuccess} onClose={() => setFormSuccess("")} />
            <TransientToast kind="error" message={formError} onClose={() => setFormError("")} />
          </form>
          <div className="managed-list">{games.length === 0 ? <div className="empty-table"><strong>暂无游戏区服</strong><span>请在左侧新增。</span></div> : games.map((game) => (
            <article className="managed-account" key={game.id ?? `${game.name}/${game.serverName}`}>
              <div><strong>{game.name} / {game.serverName}</strong><span>{game.serverAccount ? "已配置服务端账号" : "使用默认服务端账号"}</span></div>
              <div className="tag-row"><small>{game.serverUrl}</small><small>{game.iconUrl ? "已上传Icon" : "未上传Icon"}</small>{game.backgroundUrl && <small>已上传背景</small>}</div>
              <div className="managed-actions"><button onClick={() => beginEdit(game)} type="button"><UserCog size={14} />编辑</button><button onClick={async () => {
                if (!game.id) return;
                try {
                  setFormError("");
                  setFormSuccess("");
                  await onDelete(game.id);
                  setFormSuccess("区服删除成功");
                } catch (error) {
                  setFormSuccess("");
                  setFormError(error instanceof Error ? error.message : "删除区服失败");
                }
              }} type="button"><Trash2 size={14} />删除</button></div>
            </article>
          ))}</div>
        </div>
      </section>
    </div>
  );
}
ReactDOM.createRoot(document.getElementById("root")!).render(<App />);


