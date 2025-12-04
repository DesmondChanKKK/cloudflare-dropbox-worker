# cloudflare-dropbox-worker

轻量 Cloudflare Worker：从 Dropbox 下载 Excel（.xlsx），根据可配置的规则从表格中提取值并返回 JSON。

版本：1.0.1

## 功能概述

- 从 Dropbox 下载指定文件（通过 Dropbox Content API）。
- 支持多种提取配置来源：
  - wrangler.toml 中的 EXTRACTION_CONFIG（默认类型及多个 type 配置）
  - URL query 中传递的 `config`（type=custom）
  - 当 `type=raw` 时，使用请求 body（JSON 数组）里的一组规则进行提取
- 通过 `clientid` 参数做简单校验
- 支持 Refresh Token 获取短期 access token（如果配置了 REFRESH_TOKEN + APP_KEY + APP_SECRET）

## 部署

可通过 npx wrangler deploy 一键部署至 CloudFlare

## 配置（wrangler / 环境）

需要通过 `wrangler secret put` 提供以下变量：

- DROPBOX_ACCESS_TOKEN （可选，短期/静态 Token）
- DROPBOX_REFRESH_TOKEN （可选，长期刷新 Token）
- DROPBOX_APP_KEY
- DROPBOX_APP_SECRET

## 请求参数

filename (required) — 要从 Dropbox 下载的文件名（例如 invoice.xlsx）
folder (optional) — Dropbox 中的目录（可以带或不带前导 /）
type (optional) — 提取规则来源，默认 default。可选：
default：使用 EXTRACTION_CONFIG 中的 default 规则
<typeName>：使用 EXTRACTION_CONFIG 中对应的类型（例如 example_type）
custom：使用 URL query 中的 config 参数（必须是 JSON 数组）
raw：使用请求 BODY 中的 JSON 数组（单组配置，仅在 type=raw 时生效）
config (optional) — 当 type=custom 时，URL 编码的 JSON 数组，定义规则
clientid (required) — 必须等于部署时的 DROPBOX_APP_KEY（用于简单授权校验）

规则（单条）格式（数组元素）：

key: 输出字段名（string）
keywords: 用来在第一列/第二列文本中匹配的关键字数组（array of strings）
colIndex: 从哪一列取值（0-based，number）
示例规则：
[
{ "key": "hardware_total", "keywords": ["hardware","小计"], "colIndex": 3 }
]

## 请求示例

1. 使用默认配置（POST）
   使用 custom（在 URL 中传 config）

curl "https://<your-worker-domain>/?filename=invoice.xlsx&folder=invoices&type=default&clientid=YOUR_APP_KEY"

2. 使用 custom（在 URL 中传 config）
   注意：config 需 URL 编码

curl "https://<your-worker-domain>/?filename=invoice.xlsx&type=custom&clientid=YOUR_APP_KEY&config=%5B%7B%22key%22%3A%22hardware_total%22%2C%22keywords%22%3A%5B%22hardware%22%2C%22 小计%22%5D%2C%22colIndex%22%3A3%7D%5D"

3.使用 raw（将单组规则放在请求 BODY）——适合把配置放在请求体而不是 wrangler.toml

curl -X POST "https://<your-worker-domain>/?filename=invoice.xlsx&type=raw&clientid=YOUR_APP_KEY" \
 -H "Content-Type: application/json" \
 -d '[
{ "key": "hardware_total", "keywords": ["hardware", "小计"], "colIndex": 3 },
{ "key": "service_total", "keywords": ["service", "小计"], "colIndex": 3 }
]'

## 返回值说明

成功返回 JSON，通常包含 version 字段和按规则提取到的键值。例如：
{
"version": "1.0.1",
"hardware_total": 123.45,
"service_total": 67.89,
"grand_total": 191.34
}

错误返回示例：

- 400 Missing filename 或 Body/config 格式错误
- 401 Invalid clientid
- 500 Dropbox API 或 Token 问题
