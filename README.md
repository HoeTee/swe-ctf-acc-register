# SWE CTF Account Register

从报名表批量创建 GZCTF 用户账号。

规则：

```text
用户名 = 手机号 + @zjrcu.com
邮箱   = 手机号 + @zjrcu.com
密码   = 手机号后四位
```

例如：

```text
手机号: 15958153463
用户名: 15958153463@zjrcu.com
密码:   3463
```

## 文件

```text
register_accounts.py      主脚本
config.example.yaml       配置模板
requirements.txt          Python 依赖
```

`config.yaml`、Excel 文件、执行结果文件不会提交到 git。

## 安装

```powershell
python -m pip install -r requirements.txt
```

## 配置

复制配置模板：

```powershell
Copy-Item config.example.yaml config.yaml
```

推荐用环境变量保存管理员账号密码：

```powershell
$env:GZCTF_ADMIN_USERNAME="admin"
$env:GZCTF_ADMIN_PASSWORD="你的管理员密码"
```

`config.yaml` 中对应字段：

```yaml
gzctf:
  clusters:
    - name: "cluster-117"
      base_url: "http://100.99.32.117:8080"
      admin_username_env: "GZCTF_ADMIN_USERNAME"
      admin_password_env: "GZCTF_ADMIN_PASSWORD"
```

## 检查报名表

只读取 Excel，不调用 GZCTF：

```powershell
python register_accounts.py --config config.yaml --check-excel
```

如果要在检查结果里显示生成的密码：

```powershell
python register_accounts.py --config config.yaml --check-excel --print-passwords
```

## Dry Run

生成将要注册的账号列表，但不调用 GZCTF：

```powershell
python register_accounts.py --config config.yaml --dry-run
```

只看前 5 个：

```powershell
python register_accounts.py --config config.yaml --dry-run --limit 5
```

## 正式注册

```powershell
python register_accounts.py --config config.yaml
```

执行流程：

```text
1. 读取 Excel
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
8. 写入 account-register-result.json
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

默认输出到：

```text
account-register-result.json
```

每条记录包含：

```json
{
  "name": "常云凡",
  "phone": "15958153463",
  "username": "15958153463@zjrcu.com",
  "email": "15958153463@zjrcu.com",
  "password": "3463",
  "status": "created"
}
```

可能状态：

```text
created
already_exists
failed
planned
```

## 注意

这个脚本只注册用户账号。它不负责：

```text
创建队伍
修改队伍名
加入比赛
审批参赛
```

如果 GZCTF 注册后不会自动完成这些比赛关系，还需要额外脚本或平台配置处理。
