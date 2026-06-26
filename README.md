# SWE CTF Account Register

从报名表批量创建 GZCTF 用户账号。支持多个 GZCTF 集群，并且每个集群可以绑定自己的报名表。

账号规则：

```text
用户名 = 手机号 + @zjrcu.com
邮箱   = 手机号 + @zjrcu.com
密码   = 手机号后四位
```

例如：

```text
手机号 15958153463
用户名 15958153463@zjrcu.com
密码   3463
```

## 文件

```text
register_accounts.py      主脚本
config.example.yaml       配置模板
requirements.txt          Python 依赖
pyproject.toml            CLI 入口
```

不会提交到 git 的本地文件：

```text
config.yaml
*.xlsx
account-register-result.json
account-register-result.xlsx
```

## 安装

```powershell
python -m pip install -e .
```

如果你已经执行过 `pip install -e .`，普通代码修改会直接生效；只有新增或改名 CLI 命令时才需要重新安装。

## 配置

复制配置模板：

```powershell
Copy-Item config.example.yaml config.yaml
```

每个集群单独配置报名表：

```yaml
registration_defaults:
  name_column: "姓名"
  unit_column: "单位"
  department_column: "部门"
  phone_column: "联系电话"

account:
  email_domain: "zjrcu.com"
  password: "phone_last4"

gzctf:
  clusters:
    - name: "cluster-117"
      base_url: "http://100.99.32.117:8080"
      admin_username_env: "GZCTF_ADMIN_USERNAME_117"
      admin_password_env: "GZCTF_ADMIN_PASSWORD_117"
      config_update_method: "auto"
      registration:
        xlsx_path: "C:/Users/18014/OneDrive/Desktop/AI代码马拉松报名信息.xlsx"
        sheet_name: "6.24去除重复项"

registration_settings:
  allowRegister: true
  activeOnRegister: true
  useCaptcha: false
  emailConfirmationRequired: false

output:
  result_file: "account-register-result.json"
  xlsx_file: "account-register-result.xlsx"
```

多个集群就继续在 `gzctf.clusters` 下面增加条目。每个集群可以有不同的 `base_url`、管理员凭据环境变量、报名表路径和 sheet 名称。

## 管理员账号

推荐用环境变量保存管理员账号密码：

```powershell
$env:GZCTF_ADMIN_USERNAME_117 = "admin"
$env:GZCTF_ADMIN_PASSWORD_117 = "你的管理员密码"
```

PowerShell 里字符串必须加引号，否则会被当成命令。

## 检查报名表

只读取 Excel，不调用 GZCTF：

```powershell
swe-ctf-acc-register check-excel --config config.yaml
```

显示生成的密码样例：

```powershell
swe-ctf-acc-register check-excel --config config.yaml --print-passwords
```

## Dry Run

生成将要注册的账号列表，但不调用 GZCTF：

```powershell
swe-ctf-acc-register dry-run --config config.yaml
```

只看每个集群前 5 个：

```powershell
swe-ctf-acc-register dry-run --config config.yaml --limit 5
```

Dry run 也会生成结果文件：

```text
account-register-result.json
account-register-result.xlsx
```

## 正式注册

先单条验证：

```powershell
swe-ctf-acc-register register --config config.yaml --limit 1 --timeout-seconds 60 --progress-every 1
```

确认单个账号成功后，再扩大批量：

```powershell
swe-ctf-acc-register register --config config.yaml --timeout-seconds 60 --progress-every 20
```

如果批量失败，查看前几个失败详情：

```powershell
swe-ctf-acc-register register --config config.yaml --limit 5 --timeout-seconds 60 --progress-every 1 --failure-details 5
```

脚本会在每处理一个账号后增量写入：

```text
account-register-result.json
account-register-result.xlsx
```

所以即使中途 Ctrl+C 或网络异常，仍然可以打开结果文件查看已经处理到哪一条、HTTP 状态码和 GZCTF 返回内容。

执行流程：

```text
1. 按集群读取各自报名表
2. 生成 userName/email/password
3. 登录每个 GZCTF 集群管理员账号
4. GET /api/admin/config 备份当前平台配置
5. 临时设置：
   allowRegister: true
   activeOnRegister: true
   useCaptcha: false
   emailConfirmationRequired: false
6. POST /api/account/register 批量注册
7. 默认恢复原 /api/admin/config
8. 写入 JSON 和 Excel 结果表
```

平台配置更新方法默认自动尝试：

```text
PUT /api/admin/config
PATCH /api/admin/config
POST /api/admin/config
```

如果你的 GZCTF 版本只支持固定方法，可以在 cluster 下配置：

```yaml
config_update_method: "put"
```

## 输出

JSON：

```text
account-register-result.json
```

Excel：

```text
account-register-result.xlsx
```

Excel 包含三个 sheet：

```text
summary    每个集群注册汇总
accounts   每个用户账号、密码、状态、HTTP 状态码和返回消息
warnings   报名表问题或跳过原因
```

可能状态：

```text
planned
created
already_exists
failed
```

## 排障

如果控制台持续显示 `failed`，先跑：

```powershell
swe-ctf-acc-register register --config config.yaml --limit 1 --timeout-seconds 60 --progress-every 1 --failure-details 1
```

然后查看：

```powershell
Get-Content .\account-register-result.json -Raw
```

重点看第一条账号的：

```text
status
http_status
message
```

这些字段就是 GZCTF 返回的真实失败原因。
