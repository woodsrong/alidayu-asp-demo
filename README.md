# 阿里大鱼ASP调用示例
- 本demo参考[nodejs demo](https://github.com/xiaoshan5733/alidayu-node)完成
- 重点可参考sign的生成方法

> 注意:
>
> - 本demo短信签名暂时只支持4个以内汉字,待优化(目前JSONObject对象添加超过4个汉字报错)

## 代码示例

```asp
<!--#include file="alidayu-sdk/index.asp"-->

<%
'调用阿里大鱼SDK发送短信示例

Response.charset = "utf-8"
Response.Write("start <br>")

Dim mApp
Dim mParams

Set mApp = New AlidayuApp
Set mParams = New JSONobject

'修改为自己的app key和app secret，以下为测试key和secret，已失效
mApp.config "23275748", "bdfc910d42a7b421400c4b1c9c2d65c9"

'配置短信签名
mParams.add "sms_free_sign_name", "大鱼测试"
'配置接受短信手机号码
mParams.add "rec_num", "18688705716"
'配置短信模板
mParams.add "sms_template_code", "SMS_10370738"

mApp.smsSend(mParams)
%>
```

## API说明

### AlidayuApp
- Type: `Class`
- Description: 阿里大鱼app类

#### config
- Type: `Function`
- Description: 配置app key & secret

#### getSign
- Type: `Function`
- Description: 获取请求签名参数

#### request
- Type: `Function`
- Description: 请求mtop服务

#### smsSend
- Type: `Function`
- Description: 短信发送
- `params`
  - `{string} sms_free_sign_name` 短信签名 传入的短信签名必须是在阿里大鱼“管理中心-短信签名管理”中的可用签名。
  - `{json} sms_param` 短信模板变量 传参规则{"key":"value"}，key的名字须和申请模板中的变量名一致，多个变量之间以逗号隔开
  - `{string} rec_num` 短信接收号码 支持单个或多个手机号码，传入号码为11位手机号码，不能加0或+86。群发短信需传入多个号码，以英文逗号分隔，一次调用最多传入200个号码
  - `{string} sms_template_code` 短信模板ID 传入的模板必须是在阿里大鱼“管理中心-短信模板管理”中的可用模板。示例：SMS_585014
  
#### smsQuery
- Type: `Function`
- Description: 短信发送记录查询
- `params`
  - `{string} rec_num` 短信接收号码
  - `{string} query_date` 短信发送日期，支持近30天记录查询，格式yyyyMMdd
  - `{number} current_page` 分页参数,页码
  - `{number} page_size` 分页参数，每页数量。最大值100
  
#### voiceSinglecall
- Type: `Function`
- Description: 语音通知
- `params`
  - `{string} called_num` 被叫号码，支持国内手机号与固话号码,格式如下057188773344,13911112222,4001112222,95500
  - `{string} called_show_num` 被叫号显，传入的显示号码必须是阿里大鱼“管理中心-号码管理”中申请通过的号码
  - `{string} voice_code` 语音文件ID，传入的语音文件必须是在阿里大鱼“管理中心-语音文件管理”中的可用语音文件
  
#### ttsSinglecall
- Type: `Function`
- Description: 文本转语音通知
- `params`
  - `{json} tts_param` 文本转语音（TTS）模板变量，传参规则{"key"："value"}，key的名字须和TTS模板中的变量名一致，多个变量之间以逗号隔开
  - `{string} called_num` 被叫号码，支持国内手机号与固话号码,格式如下057188773344,13911112222,4001112222,95500
  - `{string} called_show_num` 被叫号显，传入的显示号码必须是阿里大鱼“管理中心-号码管理”中申请或购买的号码
  - `{string} tts_code` TTS模板ID，传入的模板必须是在阿里大鱼“管理中心-语音TTS模板管理”中的可用模板
  
#### voiceDoublecall
- Type: `Function`
- Description: 语音双呼
- `params`
  - `{string} caller_num` 主叫号码，支持国内手机号与固话号码,格式如下057188773344,13911112222,4001112222,95500
  - `{string} caller_show_num` 主叫号码侧的号码显示，传入的显示号码必须是阿里大鱼“管理中心-号码管理”中申请通过的号码。显示号码格式如下057188773344，4001112222，95500
  - `{string} called_num` 被叫号码，支持国内手机号与固话号码,格式如下057188773344,13911112222,4001112222,95500
  - `{string} called_show_num` 被叫号码侧的号码显示，传入的显示号码可以是阿里大鱼“管理中心-号码管理”中申请通过的号码。显示号码格式如下057188773344，4001112222，95500。显示号码也可以为主叫号码。
  
## 常见问题
- 如何生成sign签名?
  - 官方说明文档: <http://open.taobao.com/doc2/detail.htm?articleId=101617&docType=1&treeId=1#s4>
  - 最核心是md5加密算法,必须支持中文,推荐使用本demo中的修改版[hmac-md5.asp](http://www.yiit.cn/plugin/asp-hmac-md5-function.html)