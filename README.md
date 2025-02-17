# 036 WPS多维表格开发能力体系

本页内容

- 在线脚本AirScript
- 内嵌SDK
- WPS多维表格插件
- 能力体系概览图

# WPS多维表格能力体系 ​

WPS多维表格提供丰富的API给用户进行功能开发，用户可以借助它扩展多维表的功能。 WPS多维表格API的开放能力可以支持多种场景的应用，包括在线脚本AirScript、内嵌使用SDK、WPS多维表格插件

## 在线脚本AirScript ​

在线脚本AirScript给用户提供撰写 JavaScript 脚本的能力，使WPS多维表格能够执行自定义自动化任务。脚本是运行在服务端环境，没有前端界面，无法定义前端的界面的操作，所有的API调用都是同步调用。更多信息查看[在线脚本AirScript](/documents/app-integration-dev/guide/dbsheet/AirScript/AirScript-instro.html)

## 内嵌SDK ​

SDK提供了丰富的 **API** 对各类文档进行操作，通过使用 SDK，网页开发者可以自定义文档界面的元素、操作文档的内容、监听文档事件等操作，SDK 为用户提供了优质的在线文档体验。更多信息查看[内嵌使用SDK](/documents/app-integration-dev/guide/dbsheet/Weboffice/weboffice-instro.html)

## WPS多维表格插件 ​

WPS多维表格插件提供自定义仪表盘、视图、记录卡片的能力，可以对应用、同步表、自动化流程进行扩展。

## 能力体系概览图 ​

![WPS多维表格能力体系](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/WPSdbopen.DvCd2gVj.jpg "WPS多维表格能力体系")

# 037 在线脚本AirScript / 简介

本页内容

- AirScript 能做什么​
- 如何使用 AirScript

# 在线脚本AirScript 简介 ​

AirScript 是一个简单快速的轻量级脚本应用开发平台，给用户提供撰写 JavaScript 脚本的能力，使WPS多维表格能够执行自定义自动化任务。

## AirScript 能做什么​ ​

AirScript 目前主要为WPS多维表格打造二开平台，通过编程的方式，提供对表格数据的增删查改、单元格式修改、属性设置等能力。

工具优势​

- 无需搭建本地环境，直接在文档内进行脚本云开发。
- 内置定制化的全局 Application 对象，编辑器智能提示，开发、调试、运行一条龙服务。
- 同步获取属性，同步执行方法，减少传统的异步调用带来的心智负担。
- 得益于集成化开发环境，无论是创建定时任务，还是批量处理数据，亦或是自动化生成文档，开发者可以在这里尽情发挥自己的想象力。

## 如何使用 AirScript ​

- 打开WPS多维表格，点击「脚本」- 「JS脚本」下的 「新建脚本」，点击即可调起 AirScript 编辑器，如图： ![alt text](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/AirScript1.Du175hNg.png)![alt text](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/AirScript2.BMqsgwV7.png)
- 自动化流程中，点击「执行以下操作」 - 「执行AirScript脚本」，选择脚本和配置参数。更多自动化配置请查看 [自动化配置](/documents/app-integration-dev/guide/dbsheet/AirScript/AirScript-quickstart.html#autotask)![alt text](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/AirScript3.D3_DrowZ.png)
 
# 038 在线脚本AirScript / 快速入门

本页内容

- 在AirScript 编辑器中运行
- 在自动化流程中运行

# 开始​ ​

## 在AirScript 编辑器中运行 ​

1. 在金山文档首页新建一个WPS多维表格并打开来体验AirScript。
2. 打开WPS多维表格，点击「脚本」- 「JS脚本」下的 「新建脚本」，点击即可调起 AirScript 编辑器。
3. 将下方的例子，逐个运行，查看效果来快速上手AirScript。

javascript

```javascript

      function main(){
  console.log("hello world!")
}

main()
```

javascript

```javascript
function main(){
  // 遍历并打印所有工作表的名称
  let sheets = Application.Sheets
  for (let i = 0; i < sheets.Count; i++) {
      let sheet = sheets.Item(i + 1)
      console.log(sheet.Name) // 打印每个工作表的名称
  }
  const sheet = Application.Selection.GetActiveSheet()
  // 打印当前激活Sheet的名称
  console.log(sheet.name)

  // 打印单元格内容
  console.log(Application.ActiveView.RecordRange(1,1).Text)

  // 修改单元格内容
  Application.ActiveView.RecordRange(1,1).Value = 2;

  // 打印修改后的单元格内容
  console.log(Application.ActiveView.RecordRange(1,1).Text)
}

main()
```

> 
> 更多例子请查看 [AirScript脚本经典案例](/documents/app-integration-dev/guide/dbsheet/AirScript/AirScript-demo.html)

## 在自动化流程中运行 ​

1. 进入一个自动化流程的配置，在步骤中选择「执行AirScript脚本」，选择脚本和配置参数，「脚本入参」中配置传递给脚本的参数：

![alt text](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/airscript-quickstart1.aSmejAJq.png) 2. 在脚本中获取到传递的参数 进行开发使用：

![alt text](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/airscript-quickstart2.BdVIUnXK.png)

> 
> 更多自动化请查看 [执行AirScript脚本操作使用指南](https://kdocs.cn/l/cdQOqc6TZuMk)
 
# 039 在线脚本AirScript / 脚本令牌 / 简介

本页内容

# 脚本令牌（APIToken）​ ​

开发者通过 AirScript 编辑器编写的脚本，可以直接在编辑器内运行，也可以粘贴链接在单元格运行，或者是通过定时任务面板自动运行。

但是上述几种运行方式均集成在我们的平台方，如果开发者希望在自身的业务系统内使用到 AirScript 的能力，则需要借助我们的脚本令牌。

## 什么是脚本令牌？​ ​

脚本令牌即 APIToken，是我们为外部系统引入 AirScript 能力而专门设计的。通过脚本令牌，您可以轻松使用到金山文档 AirScript 提供的能力，执行脚本获取文档数据或者是写入文档内容。

![alt text](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/apitoken1.CxYSWxvN.png)

## 如何创建脚本令牌？​ ​

1. 进入WPS多维表格后，打开脚本编辑器，在工具栏点击【脚本令牌】按钮
2. 如果之前未创建过脚本令牌，会提示脚本令牌创建所需要注意的点，勾选【我已知晓】，然后点击【创建脚本令牌】即可
3. 为保证一定的安全，如果您还未进行过实名认证，需要先完成实名认证流程
4. 创建成功后即可获取到您的个人脚本令牌，复制令牌信息然后妥善保存

![alt text](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/apitoken2.AWNObwRN.png)

## 如何使用脚本令牌？​ ​

脚本令牌是外部执行脚本的凭证，在您成功生成自己的脚本令牌后，便可以开始着手使用脚本令牌进行脚本的执行调用。

首先，打开脚本编辑器，在侧边栏任意一个文档脚本的更多菜单里复制脚本的 webhook 链接。

![alt text](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/apitoken3.CQV0-9v7.png)

查看复制到的链接的内容如下所示：

`https://www.kdocs.cn/api/v3/ide/file/caEkI6K5RDG2/script/V2-5rSiBiN7y5xdOd5x5ZYI2r/sync_task`

链接内已拼接好了当前脚本的脚本 ID 和所在文档的文件 ID，接下来请求该链接即可读取和编辑本人相应的文档，注意调用的时候必须设置请求头`AirScript-Token`，值为您的脚本令牌，更详细的说明请参阅[接口说明](/documents/app-integration-dev/guide/dbsheet/AirScript/AirScript-apitoken-api.html)。

这里假设目标脚本的代码如下所示，将单元格的值修改为AirScript，并返回一个对象:

javascript

```javascript
Application.ActiveView.RecordRange(1,1).Value = 'AirScript';
return {
  name: '金小朦',
  age: 17
}
```

通过脚本令牌和 webhook 我们构造了一个 http 请求，如下所示：

shell

```shell
curl --request POST \
	--url https://www.kdocs.cn/api/v3/ide/file/caEkI6K5RDG2/script/V2-5rSiBiN7y5xdOd5x5ZYI2r/sync_task \
	--header 'AirScript-Token: xxx' \
	--header 'Content-Type: application/json' \
	--data '{"Context":{"argv":{}}}'
```

如果请求成功，将会返回如下数据，其内容主要包含脚本运行的日志信息和在代码中 return 的数据，如果您的脚本代码书写有误，相应的报错信息也会在日志中有所体现。

javascript

```javascript
{
  "data": {
    "logs": [
      {
        "filename": "<system>",
        "timestamp": "12:03:20.711",
        "unix_time": 1691726600711,
        "level": "info",
        "args": ["脚本环境初始化..."]
      },
      {
        "filename": "<system>",
        "timestamp": "12:03:22.129",
        "unix_time": 1691726602129,
        "level": "info",
        "args": ["已开始执行"]
      },
      {
        "filename": "<system>",
        "timestamp": "12:03:22.312",
        "unix_time": 1691726602312,
        "level": "info",
        "args": ["执行完毕"]
      }
    ],
    "result": {
      "age": 17,
      "name": "金小朦"
    }
  },
  "error": "",
  "status": "finished"
}
```

## 注意事项​ ​

1. 由于脚本令牌允许第三方访问到平台的服务端资源，为提高一定的安全性，我们需要您完成实名认证（已认证可忽略）
2. 脚本令牌，是外部执行脚本的凭证，属于个人隐私信息，通过脚本令牌配合脚本 webhook，可读取和编辑本人相应的文件，需妥善管理，请勿对外传播
3. 脚本令牌与用户绑定，每个用户最多有且仅有一个脚本令牌，创建新的令牌时，需要先对老令牌进行删除（重新创建的脚本令牌需与原令牌不同）
4. 脚本令牌默认 180 天到期，用户可在创建时手动进行延期，不做限制，可多次延期
 
# 040 在线脚本AirScript / 脚本令牌 / 接口说明

本页内容

# 接口说明​ ​

成功生成脚本令牌后，就可以通过 HTTP 接口执行脚本了，我们提供了同步执行和异步执行两种脚本执行接口供开发者使用。

相较而言，前者使用更简单，接口调用后会直接返回执行结果，适用于执行耗时一般的场景；而后者则略微复杂一点，接口调用后不会返回最终的执行的结果，但会立即返回一个`task_id`，您需要根据此`task_id`轮询脚本执行的日志，而无需同步等待结果阻塞业务流程，该接口适用于执行耗时比较大的场景。

无论使用您使用哪个接口，都必须先获取到文件 ID 和脚本 ID，请先进入脚本编辑器，在侧边栏列表的更多菜单里复制 webhook 链接即可。

## 同步执行脚本​ ​

`POST /api/v3/ide/file/:file_id/script/:script_id/sync_task`

### Header 参数 ​

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| Content-Type | 是 | string | `application/json` |
| AirScript-Token | 是 | string | 传入您通过 AirScript 编辑器生成的脚本令牌（APIToken） |

### path 参数 ​

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| script / id | 是 | string | 脚本的 ID |
| file / id | 是 | string | 运行脚本的文件 ID |

### body 参数 ​

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| Context | 是 | Object | 运行时的上下文参数 |
| Context.argv | 否 | Object | 传入的上下文参数对象，比如传入`{name: 'xiaomeng', age: 18}`，在 AS 代码中可通过`Context.argv.name`获取到传入的值 |
| Context.sheet / name | 否 | string | db,et,ksheet 运行时所在表名 |
| Context.range | 否 | string | et,ksheet 运行时所在区域，例如`$B$156` |
| Context.link / from | 否 | string | et,ksheet 点击超链接所在单元格 |
| Context.db / active / view | 否 | string | db 运行时所在 view 名 |
| Context.db / selection | 否 | string | db 运行时所在选区 |

### 返回参数 ​

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| data | 是 | Object | 任务执行数据对象 |
| data.result | 是 | string | 任务执行返回的数据 |
| data.logs | 是 | Array | 任务执行日志 |
| data.logs[i].filename | 是 | string | 执行文件的名称 |
| data.logs[i].timestamp | 是 | string | 执行时间 |
| data.logs[i].unix / time | 是 | number | 执行 unix 时间戳 |
| data.logs[i].level | 是 | string | 日志级别 |
| data.logs[i].args | 是 | string[] | 日志打印参数 |
| status | 是 | string | 任务是否执行完毕 |
| error | 是 | string | 任务执行错误信息 |
| error / details | 是 | Object | 错误信息详情对象 |
| error / details.name | 否 | string | 错误信息名称 |
| error / details.msg | 否 | string | 错误信息 |
| error / details.stack | 否 | string[] | 错误信息栈 |
| error / details.unix / time | 否 | number | 错误信息 unix 时间 |

### 请求示例 ​

shell

```shell
curl --request POST \
	--url https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/sync_task \
	--header 'AirScript-Token: xxx' \
	--header 'Content-Type: application/json' \
	--data '{"Context":{"argv":{},"sheet_name":"表名"}}'
```

### 返回示例 ​

javascript

```javascript
{
  "data": {
    "logs": [
      {
        "filename": "<system>",
        "timestamp": "16:44:08.271",
        "unix_time": 1690274648271,
        "level": "info",
        "args": ["脚本环境初始化..."]
      },
      {
        "filename": "<system>",
        "timestamp": "16:44:08.953",
        "unix_time": 1690274648953,
        "level": "info",
        "args": ["已开始执行"]
      },
      {
        "filename": "未命名脚本.js:1:9",
        "timestamp": "16:44:08.968",
        "unix_time": 1690274648968,
        "level": "info",
        "args": ["打印参数A：111"]
      },
      {
        "filename": "<system>",
        "timestamp": "16:44:08.969",
        "unix_time": 1690274648969,
        "level": "info",
        "args": ["执行完毕"]
      }
    ],
    "result": "[Undefined]"
  },
  "error": "",
  "status": "finished"
}
```

## 异步执行脚本 ​

`POST /api/v3/ide/file/:file_id/script/:script_id/task`

### Header 参数 ​

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| Content-Type | 是 | string | `application/json` |
| AirScript-Token | 是 | string | 传入您通过 AirScript 编辑器生成的脚本令牌（APIToken） |

### path 参数 ​

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| script / id | 是 | string | 脚本的 ID |
| file / id | 是 | string | 运行脚本的文件 ID |

### body 参数 ​

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| Context | 是 | Object | 运行时的上下文参数 |
| Context.argv | 否 | Object | 传入的上下文参数对象，比如传入`{name: 'xiaomeng', age: 18}`，在 AS 代码中可通过`Context.argv.name`获取到传入的值 |
| Context.sheet / name | 否 | string | db,et,ksheet 运行时所在表名 |
| Context.range | 否 | string | et,ksheet 运行时所在区域，例如`$B$156` |
| Context.link / from | 否 | string | et,ksheet 点击超链接所在单元格 |
| Context.db / active / view | 否 | string | db 运行时所在 view 名 |
| Context.db / selection | 否 | string | db 运行时所在选区 |

### 返回参数 ​

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| task / id | 是 | string | 运行的任务 Id，用于轮循运行结果 |
| task / type | 是 | string | 任务类型 |

### 请求示例 ​

shell

```shell
curl --request POST \
	--url https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/task \
	--header 'AirScript-Token: xxx' \
	--header 'Content-Type: application/json' \
	--data '{"Context":{"argv":{},"sheet_name":"表名"}}'
```

### 返回示例 ​

javascript

```javascript
{
  "data": {
    "task_id": "GN/KU3B3BG84MdCjraN5mukx0Rt5Sp1eJ9k2qClmcaOkkF3PUVNDOYPY7Kz4aQMXSvXn9N08QabldRKjPfzii87fuGYydIuK2la2HMfcxmGK1Pf4WcPEflb5xOOkQQEo8fmEbzcobhurYg=="
  },
  "task_id": "GN/KU3B3BG84MdCjraN5mukx0Rt5Sp1eJ9k2qClmcaOkkF3PUVNDOYPY7Kz4aQMXSvXn9N08QabldRKjPfzii87fuGYydIuK2la2HMfcxmGK1Pf4WcPEflb5xOOkQQEo8fmEbzcobhurYg==",
  "task_type": "open_air_script"
}
```

## 获取任务运行情况 ​

`GET /api/v3/script/task`

### query 参数 ​

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| task / id | 是 | string | 执行异步任务时返回的 ID |

> 
> 任务ID为query参数，拼接时请注意先编码下，比如`encodeURIComponent(task_id)`

### 返回参数 ​

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| data | 是 | Object | 任务执行数据对象 |
| data.result | 是 | string | 任务执行返回的数据 |
| data.logs | 是 | Array | 任务执行日志 |
| data.logs[i].filename | 是 | string | 执行文件的名称 |
| data.logs[i].timestamp | 是 | string | 执行时间 |
| data.logs[i].unix / time | 是 | number | 执行 unix 时间戳 |
| data.logs[i].level | 是 | string | 日志级别 |
| data.logs[i].args | 是 | string[] | 日志打印参数 |
| status | 是 | string | 任务是否执行完毕 |
| error | 是 | string | 任务执行错误信息 |
| error / details | 否 | object | 错误信息详情对象 |
| error / details.name | 否 | string | 错误信息名称 |
| error / details.msg | 否 | string | 错误信息 |
| error / details.stack | 否 | string[] | 错误信息栈 |
| error / details.unix / time | 否 | number | 错误信息 unix 时间 |

### 请求示例 ​

shell

```shell
curl --request GET \
	--url https://www.kdocs.cn/api/v3/script/task
```

### 返回示例 ​

json

```json
{
  "data": {
    "logs": [
      {
        "filename": "<system>",
        "timestamp": "17:05:16.164",
        "unix_time": 1692090316164,
        "level": "info",
        "args": ["脚本环境初始化..."]
      }
    ],
    "result": null
  },
  "error": "Unexpected token (1:91)",
  "error_details": {
    "name": "SyntaxError",
    "msg": "Unexpected token (1:91)",
    "stack": ["    at 未命名脚本.js:1:91"],
    "unix_time": 1692090318372
  },
  "status": "finished"
}
```
 
# 041 在线脚本AirScript / 脚本令牌 / 应用场景

本页内容

- 1. 私密信息查询​
- 2. 电商数据同步​
- 3. RPA 数据同步​
- 4. 简易数据库​
- 5. 作为数据工具使用​

# 应用场景​ ​

在脚本令牌的加持下，WPS多维表格的强大能力得到完美释放，开发者能以一种更为高效和准确的方式执行任务。无论是网页爬取、数据分析，还是自动化流程，我都可以借助脚本令牌来完成。

如下为我们根据实际场景写的一些示例和说明，希望能给开发者一定的启发。

## 1. 私密信息查询​ ​

现代社会越来越重视个人的隐私。这种趋势在很多方面都有所体现。

首先，在教育领域，许多学校开始加强对学生在校期间的信息保护，禁止将学生的个人信息出售或分享给他人。这包括学生的家庭信息、教育记录、考试成绩等。

其次，在医疗保健领域，病人的隐私保护成为了一个重要的问题。医生和其他医疗工作者需要遵守严格的隐私规定，确保病人的个人信息不会被泄露。

此外，在社交媒体领域，许多平台也开始加强对用户信息的保护。他们采取了更严格的数据安全措施，以确保用户的数据不会被泄露或滥用。

如下，我们将展示一个学生成绩查询的在线示例，你可以在线体验利用脚本令牌实现的私密信息查询的功能。现有学生成绩表如下图所示：

![alt text](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/airscript-demo1.CieIZ_3a.png)

你可以根据上表提供的学生信息，输入对应学生的学号和姓名后，即可查询到对应学生的成绩，下面有脚本的示例。

> 
> 注意 真实使用场景中，会要求输入密码或者手机验证等。

![alt text](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/airscript-demo2.DejAWVmS.png)

javascript

```javascript
function main() {
  // 获取学生的学号
  const student_id = '1'
  // 获取学生的姓名
  const student_name = '金小妹'
  // 获取记录区域
  const rg = Application.Sheets(1).Views(1).RecordRange
  // 获取记录区域的行数
  const rowCount = rg.Count

  // 获取匹配的学生数据
  const data = []
  for (let i = 1; i <= rowCount; i++) {
    const studentId = rg(i, "@学号").Value
    const studentName = rg(i, "@姓名").Value

    if (studentId === student_id && studentName === student_name) {
      const sex = rg(i, "@性别").Value
      const className = rg(i, "@班级").Value
      const language = rg(i, "@语文").Value
      const math = rg(i, "@数学").Value
      const english = rg(i, "@英语").Value
      
      data.push({
        id: studentId,
        name: studentName,
        sex: sex,
        className: className,
        language: language,
        math: math,
        english: english,
      })
      break
    }
  }
  // 返回匹配的学生数据
  return data
}
main()
```

## 2. 电商数据同步​ ​

电商数据同步是一个重要的环节，确保电商平台在销售和运营方面能够高效运作。在数据来源上，它可能是多个不同的系统，包括数据库、ERP系统、CRM系统等。而开发者首先要做的是手动导出或者 webhook 的形式获取到商品信息、订单信息、物流信息等数据，然后在您的个人服务器内得到这些数据，进行数据清洗和转换，最后再通过脚本令牌写入到金山文档多维表格内，完成数据的同步。

## 3. RPA 数据同步​ ​

如果开发者有自己的 RPA 平台，需要从金山文档内获取值班信息，再通过 RPA 发送到他们的工作群内。传统的实现方式很麻烦，需要先利用 AirScript 的邮件服务将内容发送到邮件里，然后新建定时任务定时执行脚本，最后通过 RPA 读取邮件内容后转发到企业微信。

有了脚本令牌后，再也不用这么辛苦的“曲线救国”了，想要什么数据，通过脚本令牌直线获取即可。

## 4. 简易数据库​ ​

一些开发者有自己的个人网站，用户量很少或者仅作为个人学习使用，直接购买云数据库成本太高不划算。

此时大家可以想想，数据库里的表是表，多维表格里的表也是表，在某种条件下，有没有可能多维表格可以平替掉云数据库？

当然有可能，使用脚本令牌您可以轻松的完成数据的增删改查，扔掉老爷车 SQL，使用 JavaScript 来进行“为所欲为”的结构化查询，快来体验一下吧。

> 
> 注意 以上仅提供一个新思路，只适用于个人学习或者非常轻量级的服务，毕竟玩归玩闹归闹，别拿数据开玩笑。

## 5. 作为数据工具使用​ ​

有时候我们需要在自己的系统内使用到表格提供的高级功能，来完成对数据的筛选和过滤操作。比如现在有这么一个场景：开发者使用影刀 RPA 进行一个网站的数据爬取，爬取完了之后存到一个 excel，对 excel 做完数据清理之后才能进行下一步操作。但是有了脚本令牌后，就可以先将数据写入金山文档表格中，然后执行 AirScript 基于规则进行一个自动清理，直接就能进行下一步操作。
 
# 042 在线脚本AirScript / 高级服务 / 简介

本页内容

# 简介 ​

借助AirScript的高级服务，开发者只需要完成较少设置，即可连接到某些公开的金山文档API。 它们的使用方式与AirScript脚本的内置函数十分相似。

AirScript在运行时会自动处理授权流程。 不过开发者必须启用高级服务，才能在脚本中使用该服务，若跳过该步骤，会因为找不到该服务而抛出**undefined**错误。

## 启用高级服务​ ​

要使用高级服务，请按以下说明操作：

1. 打开**AirScript编辑工具**弹出编辑页面。
2. 点击AirScript编辑工具上方的**服务**。
3. 点击**添加服务**。 ![alt text](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/airscript-advanced1.BNcn59rm.png)
4. 选择一项服务，然后点击**确认**。

启用高级服务后，该服务会在自动补全中显示。

## 授权流程​ ​

AirScript需要用户授权才能访问高级服务中的私密数据。

### 授予运行权限​ ​

AirScript会根据开发者编写脚本时启用高级服务的配置内容来确定授权范围 （例如访问指定文件或访问网络）。如果脚本需要授权，用户在运行脚本时会弹出授权对话框。 描述这个脚本涉及到的授权范围。

普通的代码更改并不会清空用户对脚本的授权。但如果开发者对更改了高级服务的配置（新增，修改或删除）， 那用户对脚本的授权也会清空，再次运行脚本时会重新触发授权流程。

### 取消授权​ ​

用户可以对已授权的脚本手动取消授权，请按以下说明操作

1. 打开**AirScript编辑工具**弹出编辑页面。
2. 找到脚本列表下想取消授权的脚本，点击 … 显示更多操作。
3. 点击**取消服务授权**

## 使用限制​ ​

为防止向用户提供恶意的脚本，出于安全性考虑，使用高级服务存在一些限制。

- 过于高频地使用高级服务，当出现这种情况时，脚本的运行会抛出明显的错误通知用户异常调用。
- 使用[HTTP](/documents/app-integration-dev/guide/dbsheet/AirScript/AirScript-advanced-http.html)服务时，禁止使用IP地址发起请求，禁止使用端口发起请求。
- 使用[HTTP](/documents/app-integration-dev/guide/dbsheet/AirScript/AirScript-advanced-http.html)服务时，收到内容的消息体最大为2M，超过2M会抛出错误。
- 使用[KSDrive.openFile](/documents/app-integration-dev/guide/dbsheet/AirScript/AirScript-advanced-KSDrive.html) 获得的File对象没有调用close, 就再次使用KSDrive.openFile 会报错。
 
# 043 在线脚本AirScript / 高级服务 / 网络 API

本页内容

- 快速使用
- 方法列表
- fetch(url[, options])
    - 示例
    - 参数
    - RequestOption
    - 返回值
- get(url[, options])
    - 示例
    - 参数
    - MethodRequestOption
    - 返回值
- delete(url[, options])
    - 示例
    - 参数
    - 返回值
- post(url,body[, options])
    - 示例
    - 参数
    - 返回值
- put(url,body[, options])
    - 示例
    - 参数
    - 返回值
- Response
    - 示例
    - 方法列表
- status
    - 示例
    - 返回值
- statusText
    - 示例
    - 返回值
- headers
    - 示例
    - 返回值
- text()
    - 示例
    - 返回值
- json()
    - 示例
    - 返回值
- binary()
    - 示例
    - 返回值

# 网络 API ​

AirScript 提供一个全局的 HTTP 对象，开发者可通过此对象提供的方法请求外部服务，请求成功后会同步返回服务器的响应。

该 API 的使用方式与浏览器内的 fetch()函数基本一致，对于前端开发者来说应该可以很快上手。

> 
> 在使用 HTTP 对象提供的方法发送请求之前，确保您已添加`网络API`服务，在脚本编辑器的【工具栏】-【服务】菜单内添加即可。

### 快速使用 ​

javascript

```javascript
// 发起网络请求
const resp = HTTP.fetch('https://open.iciba.com/dsapi/', {
  timeout: 2000
})
const data = resp.json()
console.log(data.note, data.content)
```

### 方法列表 ​

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| fetch(url[, options]) | Response | 发起自定义类型的网络请求 |
| get(url[, options]) | Response | 发起 GET 类型的网络请求 |
| delete(url[, options]) | Response | 发起 DELETE 类型的网络请求 |
| post(url,body[, options]) | Response | 发起 POST 类型的网络请求 |
| put(url,body[, options]) | Response | 发起 PUT 类型的网络请求 |

## fetch(url[, options]) ​

发起一个网络请求，可以自定义设置 headers 和 body。

### 示例 ​

javascript

```javascript
const resp = HTTP.fetch('https://www.kdocs.cn', {
  method: 'GET',
  timeout: 2000,
  headers: {
    'User-Agent':
      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36'
  }
})
console.log(resp.text())
```

### 参数 ​

| 名称 | 类型 | 默认值 | 必填项 | 说明 |
| --- | --- | --- | --- | --- |
| url | string |  | true | 需要访问的网络地址，只允许访问不带端口号的域名 |
| options | RequestOption | undefined | false | 一个 JavaScript 对象，可指定发起请求的可选参数，如下所示。 |

### RequestOption ​

| 名称 | 类型 | 默认值 | 必填项 | 说明 |
| --- | --- | --- | --- | --- |
| method | string | GET | false | 发起网络请求的方法，例如`GET`、`POST`、`PUT`、`DELETE`等 |
| timeout | number | 10000 | false | 发起网络请求的超时时间，单位毫秒(ms)，数据范围为 0~60000，超出范围的数据将被设为默认值 10 秒。 |
| headers | object | undefined | false | 发起网络请求的头部。例如`cookie`等 |
| body | string | undefined | false | 发起网络请求的主体内容。 |

### 返回值 ​

Response - 服务器返回的响应

## get(url[, options]) ​

发起 GET 类型的网络请求。

### 示例 ​

javascript

```javascript
const resp = HTTP.get('https://reqres.in/api/users/2')
console.log(resp.json())
```

### 参数 ​

| 名称 | 类型 | 默认值 | 必填项 | 说明 |
| --- | --- | --- | --- | --- |
| url | string |  | true | 需要访问的网络地址，只允许访问不带端口号的域名 |
| options | MethodRequestOption | undefined | false | 一个 JavaScript 对象，可指定特定请求的可选参数，如下所示。 |

### MethodRequestOption ​

| 名称 | 类型 | 默认值 | 必填项 | 说明 |
| --- | --- | --- | --- | --- |
| timeout | number | 10000 | false | 发起网络请求的超时时间，单位毫秒(ms)，数据范围为 0~60000，超出范围的数据将被设为默认值 10 秒。 |
| headers | object | undefined | false | 发起网络请求的头部。例如`cookie`等 |

### 返回值 ​

Response - 服务器返回的响应

## delete(url[, options]) ​

发起 DELETE 类型的网络请求。

### 示例 ​

javascript

```javascript
const resp = HTTP.delete('https://reqres.in/api/users/2')
console.log(resp.status)
```

### 参数 ​

| 名称 | 类型 | 默认值 | 必填项 | 说明 |
| --- | --- | --- | --- | --- |
| url | string |  | true | 需要访问的网络地址，只允许访问不带端口号的域名 |
| options | MethodRequestOption | undefined | false | 一个 JavaScript 对象，可指定特定请求的可选参数，如下所示。 |

### 返回值 ​

Response - 服务器返回的响应

## post(url,body[, options]) ​

发起 POST 类型的网络请求。

### 示例 ​

javascript

```javascript
// 发送form
const formResp = HTTP.post(
  'https://www.example.cn',
  { foo: 'bar' },
  { headers: { 'content-type': 'multipart/form-data' } }
)

//发送json
const resp = HTTP.post('https://reqres.in/api/users', {
  name: 'morpheus',
  job: 'leader'
})

console.log(resp.json())
```

### 参数 ​

| 名称 | 类型 | 默认值 | 必填项 | 说明 |
| --- | --- | --- | --- | --- |
| url | string |  | true | 需要访问的网络地址，只允许访问不带端口号的域名 |
| body | string| object |  | true | 请求体 |
| options | MethodRequestOption | undefined | false | 一个 JavaScript 对象，可指定特定请求的可选参数，如下所示。 |

### 返回值 ​

Response - 服务器返回的响应

## put(url,body[, options]) ​

发起 PUT 类型的网络请求。

### 示例 ​

javascript

```javascript
const resp = HTTP.put('https://reqres.in/api/users/200', {
  name: 'wps',
  job: 'developer'
})
console.log(resp.json())
```

### 参数 ​

| 名称 | 类型 | 默认值 | 必填项 | 说明 |
| --- | --- | --- | --- | --- |
| url | string |  | true | 需要访问的网络地址，只允许访问不带端口号的域名 |
| body | string| object |  | true | 请求体 |
| options | MethodRequestOption | undefined | false | 一个 JavaScript 对象，可指定特定请求的可选参数，如下所示。 |

### 返回值 ​

Response - 服务器返回的响应

## Response ​

HTTP 发起网络请求后返回的响应，response 是流数据，只有首次调用 text()，json()或 binary()能获取到数据

### 示例 ​

javascript

```javascript
let resp = HTTP.get('https://open.iciba.com/dsapi/')
console.log(resp.status) // 200
console.log(resp.statusText) // OK
console.log(resp.text()) // `{foo:"bar"}`
console.log(resp.json()) // {foo:"bar"}
console.log(resp.status) // [...]
```

### 方法列表 ​

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| status | number | 获取响应的 HTTP 状态码 |
| statusText | string | 获取响应的 HTTP 状态 |
| headers | object | 获取响应的 header |
| text() | string | 获取服务器返回的文本 Body |
| json() | any | 将服务器返回的 json 类型的 Body 转化为结构体 |
| binary() | [Buffer](https://nodejs.org/docs/latest-v17.x/api/buffer.html) | 获取服务器返回的二进制结构的 Body |

## status ​

获取响应的 HTTP 状态码

### 示例 ​

javascript

```javascript
const resp = HTTP.get('https://open.iciba.com/dsapi/')
console.log(resp.status) // 200
```

### 返回值 ​

number - 服务器返回响应的 HTTP 状态码

## statusText ​

获取响应的 HTTP 状态

### 示例 ​

javascript

```javascript
const resp = HTTP.get('https://open.iciba.com/dsapi/')
console.log(resp.statusText) // OK
```

### 返回值 ​

string - 服务器返回响应的 HTTP 状态

## headers ​

获取响应的 header

### 示例 ​

javascript

```javascript
let resp = HTTP.get('https://open.iciba.com/dsapi/')
console.log(resp.headers) // {"content-length":"44","content-type":"text/html; charset=utf-8"}
```

### 返回值 ​

object - 服务器返回响应的 header

## text() ​

获取服务器返回的 Body

### 示例 ​

javascript

```javascript
let resp = HTTP.get('https://open.iciba.com/dsapi/')
console.log(resp.text()) // this is an example.
```

### 返回值 ​

string - 服务器返回的响应的 Body，以文本接受并返回

## json() ​

获取服务器返回的 Body

### 示例 ​

javascript

```javascript
let resp = HTTP.get('https://open.iciba.com/dsapi/')
console.log(resp.json()) // {msg:"this is an example."}
```

### 返回值 ​

Object, Array, string, number, boolean, or null - 服务器返回的响应的 Body，以文本接受并经过 JSON.parse()后返回

## binary() ​

获取服务器返回的 Body

### 示例 ​

javascript

```javascript
let resp = HTTP.get('https://open.iciba.com/dsapi/')
console.log(resp.binary().toString('base64'))
```

### 返回值 ​

[Buffer](https://nodejs.org/docs/latest-v17.x/api/buffer.html) - 服务器返回的响应的 Body，以 Buffer 接受二进制数据并返回
 
# 044 在线脚本AirScript / 高级服务 / 云文档 API

本页内容

# 云文档 API ​

AirScript 提供全局的 KSDrive 对象，通过此对象即可轻松**查看、修改、创建和删除**您的云文档

> 
> 在使用 KSDrive 对象操作云文档时，确保您已添加`云文档API`服务，在脚本编辑器的服务菜单内添加即可。

### 快速使用 ​

js

```js
// 打开指定文档
let file = KSDrive.openFile('https://www.kdocs.cn/l/xxxxxxxxxxxx')
// 打印指定文档的A1单元格内容
console.log(file.Application.Range('A1').Text)
// 使用结束之后调用close关闭文档，否则无法再次调用KSDrive.openFile
file.close()
// 获取我的云文档下面的et，ksheet文档列表
const fileList = KSDrive.listFiles({ includeExts: ['et', 'ksheet'] })
// 打开我的云文档目录下的第一个文档
file = KSDrive.openFile(fileList.files[0])
console.log(file.Application.Range('A1').Text)
// 关闭文档
file.close()
```

### 属性列表 ​

| 属性名 | 数据类型 | 说明 |
| --- | --- | --- |
| FileType | object | 支持的文件类型集合 |

### 方法列表 ​

| 方法名 | 返回类型 | 说明 |
| --- | --- | --- |
| createFile() | string | 创建或另存一个文件 |
| openFile() | File | 额外打开一个文件 |
| listFiles() | FilesInfo | 列出某个目录下的表格文件 |

## FileType ​

云文档支持的文件类型，可用于新建文件时指定新文件的类型

### 属性说明 ​

| 属性名 | 数据类型 | 说明 |
| --- | --- | --- |
| AP | string | 智能文档 |
| KSheet | string | 智能表格 |
| ET | string | 表格 |
| DB | string | 多维表 |

## createFile() ​

创建一个新文件，也可以将一个源文件另存为新文件

### 参数 ​

| 名称 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| type | FileType | 是 | 新文件的类型 |
| createOptions | CreateOptions | 是 | 新文件的参数选项 |

### CreateOptions 对象说明 ​

| 名称 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| name | string | 是 | 新文件的文件名 |
| dirUrl | string | 否 | 新文件的文件目录 |
| source | string | 否 | 将目标文件另存为新文件 |

### 返回值 ​

url - string 新文件的 URL

### 示例 ​

js

```js
// 创建ET文件，指定保存位置
let url = KSDrive.createFile(KSDrive.FileType.ET, {
  name: 'et测试',
  dirUrl: '指定保存位置'
})
console.log(url)
// 新建DB文件
url = KSDrive.createFile(KSDrive.FileType.DB)
console.log(url)
// 新建KSheet文件
url = KSDrive.createFile(KSDrive.FileType.KSheet)
console.log(url)
// 新建AP文件
url = KSDrive.createFile(KSDrive.FileType.AP)
console.log(url)
// 文件另存
url = KSDrive.createFile(KSDrive.FileType.KSheet, {
  source: 'https://www.kdocs.cn/l/cqQwuiG2mo7E',
  name: '复制表格'
})
console.log(url)
```

## openFile() ​

额外打开一个文件，并返回一个 JavaScript 对象File。

### 示例 ​

js

```js
let file = KSDrive.openFile('https://www.kdocs.cn/l/xxxxxxxxxxxx')
console.log(file.Application.ActiveSheet.Range('A1').Text)
file.close()
```

### 参数 ​

| 名称 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| openInfo | URL / FileInfo | 是 | 打开文件的信息，可以为文件分享链接或者FileInfo |

### 返回值 ​

File - 一个 JavaScript 对象

## listFiles() ​

列出某个目录下的所有文件和对应信息

### 示例 ​

js

```js
// 遍历获取某个文件夹下的所有文件的文件名
for (let offset = 0; offset >= 0; ) {
  const list = KSDrive.listFiles({
    dirUrl: 'https://www.kdocs.cn/mine/xxxxxxxxxx',
    offset: offset,
    count: 100
  })
  for (let i = 0; i < list.files.length; i++) {
    console.log(list.files[i].fileName)
  }
  offset = list.nextOffset
}
```

### 参数 ​

| 名称 | 类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| options | object | undefined | 否 | 一个 JavaScript 对象，undefined 时获取我的云文档目录下面的文件数据，详细参数如下所示 |

### 详细参数 ​

| 参数名 | 参数类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| dirUrl | string |  | false | 目录链接，如`https://www.kdocs.cn/mine/xxxxxx`，为空时获取我的云文档目录下面的文件数据 |
| offset | number | 0 | false | 开始位置。通常由listFiles()函数返回。比如，listFiles()函数在某次检索中返回了 nextOffset 为 100，而想要获取更多文件信息，则下一次调用listFiles()函数时把 100 作为此可选参数传入。 |
| count | number | 30 | false | 文件个数 |
| includeExts | string[] |  | false | 指定文件类型,支持参数及对应关系，ksheet:"表格",et:"WPS 表格",db:"多维表",otl:"文档",wpp:"演示",wps:"WPS 文字" |

### 返回值 ​

FilesInfo - 一个 JavaScript 对象，文件信息

## File ​

打开文件函数openFile()返回的一个 JavaScript 对象。

### 属性 ​

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| Application | Application(ET/Ksheet/DBT) | 被打开文件的操作对象，目前支持 et,ksheet,dbt |
| close | Function | 关闭文件的函数，使用完 file 对象之后调用，关闭打开的文件 |

## FilesInfo ​

获取文件夹信息函数listFiles(options)返回的一个 JavaScript 对象。

### 属性 ​

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| files | FileInfo[] | 文件信息，详细参数如下所示 |
| nextOffset | number | 下一页的偏移量，可以作为listFiles(options)的参数而输出下一页文件内容，当下一页为空时，nextOffset 为-1 |

### FileInfo ​

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| fileName | string | 文件名 |
| fileId | string | 加密后的文件 id |
| createTime | number | 文件创建时间戳 |
| updateTime | number | 文件修改时间戳 |
 
# 045 在线脚本AirScript / 高级服务 / 邮件 API

本页内容

# 邮件 API ​

通过外部邮件服务发送邮件。

### 快速使用 ​

javascript

```javascript
// 登录
let mailer = SMTP.login({
    host: "smtp.example.com", // 域名
    port: 465, // 端口
    secure: true, // TLS
    username: "sender@example.com", // 账户名
    password: "Pa55W0rd" // 密码
})
// 客户端发送邮件
mailer.send({
    from: "sender@example.com", // 发件人
    to: "reciever@example.com", // 收件人
    subject: "this is subject.", // 主题
    text: "this is text.", // 文本
    html: `<p> this is html </p>` // HTML代码
})
// 支持指定昵称
mailer.send({
    from: "管理员 <admin@example.com>",
    to: "接受者 <username@example.com>",
    subject: "this is subject.",
    text: "this is text.",
    html: `<p> this is html </p>`
})
// 支持发送多个邮箱
mailer.send({
    from: "管理员 <admin@example.com>",
    to: ["username1@example.com","接受者2 <username2@example.com>"],
    subject: "this is subject.",
    text: "this is text.",
    html: `<p> this is html </p>`
})
```

## SMTP ​

### 方法列表 ​

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| login(argvs) | Mailer | 登录并返回邮件发送者 |

## login(argvs) ​

登录并返回邮件发送对象

javascript

```javascript
//  登录qq邮箱
let mailer = SMTP.login({
    host: "smtp.qq.com", // QQ 的SMTP服务器的域名
    port: 465,
    username: "1000000000@qq.com", // qq 邮箱地址
    password: "xxxxxxxxxxxx", // qq邮箱的SMTP密码，非qq密码
    secure: true
});
```

### 参数 ​

| 名称 | 类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| argvs | LoginArgvs | undefined | true | 一个JavaScript对象，用于配置SMTP的参数，如下所示 |

### LoginArgvs ​

| 名称 | 类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| host | string | undefined | true | 邮箱服务器域名 |
| port | number | undefined | true | SMTP服务端口，当host为undefined时，取默认值，默认值由secure决定，当secure是false时默认值为587，当secure是true时默认值为465。 |
| secure | boolean | undefined | true | 是否使用TLS连接服务器，在大多数情况下，如果要连接到端口465，请将此值设置为true；如果要连接到端口587或25，请将此值设置为false。 |
| username | string | undefined | true | 用于身份验证的账户名 |
| password | string | undefined | true | 用于身份验证的密码 |
| timeout | number | 10000 | false | 等待建立连接的时间，单位毫秒(ms) |

### 返回值 ​

Mailer- 邮件发送者

## Mailer ​

由 login(argvs)创建的对象，用于发送邮件

### 方法列表 ​

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| send(message) | undefined | 发送邮件 |

## send(message) ​

发送邮件

javascript

```javascript
mailer.send({
    from: ["Administrator <admin@example.com>"],
    to: ["username@example.com", "UserName <username2@example.com>"],
    subject: "this is subject.",
    text: "this is text.",
    html: `<p> this is html </p>`
})
```

### 参数 ​

| 名称 | 类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| message | messageArgvs | undefined | true | 一个JavaScript对象，要发送的邮件内容，如下所示 |

### messageArgvs ​

| 名称 | 类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| from | string | undefined | true | 发件人的电子邮箱地址 |
| to | string / string[] | undefined | true | 收件人的电子邮箱地址 |
| subject | string | undefined | true | 电子邮件的主题 |
| text | string | undefined | true | 电子邮件显示的文本 |
| html | string | undefined | false | 电子邮件的HTML代码 |
 
# 046 在线脚本AirScript / 高级服务 / 数据库 API

本页内容

# 数据库 API ​

AirScript 提供一个全局的 SQL 对象，开发者可通过此对象提供的属性和方法连接到**外部数据库**服务，连接成功后即可执行 SQL 语句，对数据进行增删改查。

> 
> 在使用 SQL 对象连接数据库之前，确保您已添加`数据库API`服务，在脚本编辑器的【工具栏】-【服务】菜单内添加即可。

### 快速使用 ​

js

```js
// 连接MySQL数据库
const connection = SQL.connect(SQL.Drivers.MySQL, {
  host: '127.0.0.1',
  username: 'root',
  password: '123456',
  database: 'mydb',
  port: 3306
})

// 执行SQL语句，查询test表的所有数据
const result1 = connection.queryAll('SELECT * FROM test')
// 打印执行结果
console.log(result1)

// 执行SQL语句，插入数据
const result2 = connection.queryAll(
  'INSERT INTO test (id,test_data) VALUES (?,?), (?,?)',
  [1, 1, 2, 2]
)
// 打印执行结果
console.log(result2)

// 关闭数据库连接
connection.close()
```

### 属性列表 ​

| 属性名 | 数据类型 | 说明 |
| --- | --- | --- |
| Drivers | object | 数据库连接驱动集 |
| Types | object | 数据库字段类型集（仅适用于 SQL server） |

### 方法列表 ​

| 方法 | 返回类型 | 说明 |
| --- | --- | --- |
| connect() | Connection | 连接目标数据库 |
| Connection.queryAll() | Result | 执行 SQL 语句 |
| Connection.close() | null | 关闭数据库连接 |

## Drivers ​

数据库驱动集，调用connect()方法连接数据库时传入对应驱动，目前仅支持 MySQL 和 SQL server 两种驱动，只读

### 属性说明 ​

| 属性名 | 数据类型 | 说明 |
| --- | --- | --- |
| MySQL | string | MySQL 数据库驱动 |
| PostgreSQL | string | PostgreSQL 数据库驱动 |
| SQLServer | string | SQL server 数据库驱动 |

## Types ​

数据库字段类型集，请注意，该类型集仅适用于 SQL server 数据库，MySQL 数据库不需要传递此值

### 属性说明 ​

#### Exact numerics ​

| 属性名 | 对应 Javascript 类型 |
| --- | --- |
| Bit | Boolean |
| TinyInt | Number |
| SmallInt | Number |
| Int | Number |
| BigInt | String |
| Numeric | Number |
| Decimal | Number |
| SmallMoney | Number |
| Money | Number |

#### Approximate numerics ​

| 属性名 | 对应 Javascript 类型 |
| --- | --- |
| Float | Number |
| Real | Number |

#### Date and Time ​

| 属性名 | 对应 Javascript 类型 |
| --- | --- |
| SmallDateTime | Date |
| DateTime | Date |
| DateTime2 | Date |
| DateTimeOffset | Date |
| Time | Date |
| Date | Date |

#### Character Strings ​

| 属性名 | 对应 Javascript 类型 |
| --- | --- |
| Char | String |
| VarChar | String |
| Text | String |

#### Unicode Strings ​

| 属性名 | 对应 Javascript 类型 |
| --- | --- |
| NChar | String |
| NVarChar | String |
| NText | String |

#### Binary Strings ​

| 属性名 | 对应 Javascript 类型 |
| --- | --- |
| Binary | Buffer |
| VarBinary | Buffer |
| Image | Buffer |

#### Other Data Types ​

| 属性名 | 对应 Javascript 类型 |
| --- | --- |
| Null | null |
| TVP | Object |
| UDT | Buffer |
| UniqueIdentifier | String |
| Variant | any |
| xml | String |

## connect() ​

连接目标数据库，目前仅支持 MySQL、PostgreSQL 和 SQL server 三种类型的数据库，连接成功后会返回数据库连接对象，可通过此对象执行 SQL 语句，程序结束之前请调用close()方法关闭数据库连接。

### 参数 ​

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| driver | Driver | null | 是 | 指定目标数据库驱动 |
| options | Options | null | 是 | 数据库连接信息 |

### options 对象说明 ​

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| host | string | null | 是 | 目标数据库主机 |
| port | number | null | 是 | 目标数据库端口 |
| username | string | null | 是 | 目标数据库连接用户名 |
| password | string | null | 是 | 目标数据库连接密码 |
| database | string | null | 是 | 目标数据库名 |

### 返回值 ​

Connection - 数据库连接对象

### 示例 ​

js

```js
// 连接MySQL数据库
const connection = SQL.connect(SQL.Drivers.MySQL, {
  host: '127.0.0.1',
  port: 3340,
  username: 'jinxiaomeng',
  password: '123',
  database: 'WPS_TEST'
})
```

## Connection.queryAll() ​

通过上述的connect()方法成功连接数据库后，会返回数据库连接对象，通过此对象即可调用 queryAll()方法执行 SQL 语句

### 参数 ​

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| sql | string | null | 是 | 要执行的 sql 语句 |
| InsertData | any[] | InsertData | null | 否 | 需要插入的数据 |

### InsertData 对象说明 ​

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| name | string | null | 是 | 插入数据的字段名 |
| value | string | null | 是 | 插入数据的值 |
| type | Types | null | 否 | 插入数据的类型，SQL server 数据库必须传递该类型 |

### 返回值 ​

Result 对象，包含受影响的行数以及返回的数据行

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| affectRowCount | number | 执行 sql 语句后受到影响的行数 |
| rows | Array | 数据行，根据实际查询的表的数据结构返回 |

### 返回示例 ​

json

```json
// 查询时的返回
{
  "affectRowCount": 0,
  "rows": [
    [
      {
        "name": "1",
        "value": 2
      }
    ]
  ]
}

// 增删改时的返回
{
  "affectRowCount": 1,
  "rows": []
}
```

### 示例 ​

js

```js
// 连接SQL server数据库
const connection = SQL.connect(SQL.Drivers.SQLServer, {
  host: 'x.x.x.x',
  username: 'x',
  password: 'x',
  database: 'x',
  port: 1433
})

// 执行sql语句，插入两条数据
const result1 = connection.queryAll(
  'INSERT INTO TestSchema.Employees (Name, Location) OUTPUT INSERTED.Id VALUES (@Name, @Location);',
  [
    {
      name: 'Name',
      type: SQL.Types.NVarChar,
      value: 'zhangsan'
    },
    {
      name: 'Location',
      type: SQL.Types.NVarChar,
      value: 'zhuhai'
    }
  ]
)

// 打印执行结果
console.log(result1)

// 执行sql语句，查询员工表
const result2 = connection.queryAll(
  'SELECT Id, Name, Location FROM TestSchema.Employees;'
)

// 打印执行结果
console.log(result2)

// 关闭数据库连接
connection.close()
```

## Connection.close() ​

关闭数据库连接，请务必在程序结束前调用此方法

### 示例 ​

js

```js
// 连接MySQL数据库
const connection = SQL.connect(SQL.Drivers.MySQL, {
  host: '127.0.0.1',
  port: 3340,
  username: 'jinxiaomeng',
  password: '123',
  database: 'WPS_TEST'
})

// do something

// 关闭数据库连接
connection.close()
```
 
# 047 在线脚本AirScript / 内置基础类型

本页内容

# 内置基础类型 ​

内置的基本数据类型、对象和函数用来帮助开发者开发，遵循**JavaScript**函数命名的标准规范，同时不能和OpenApi提供对象重名。

## 基本数据类型： ​

- [Boolean type](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#boolean_type)
- [Null type](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#null_type)
- [Undefined type](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#undefined_type)
- [Number type](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#number_type)
- [String type](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#string_type)

## 内置对象： ​

- Object
- Array
- Map
- JSON
- BigInt
- Math
- Date
- Error

## 内置函数： ​

- [isNaN()](https://developer.mozilla.org/zh-CN/docs/Web/JavaScript/Reference/Global_Objects/isNaN)
- [parseFloat()](https://developer.mozilla.org/zh-CN/docs/Web/JavaScript/Reference/Global_Objects/parseFloat)
- [parseInt()](https://developer.mozilla.org/zh-CN/docs/Web/JavaScript/Reference/Global_Objects/parseInt)
- [decodeURI()](https://developer.mozilla.org/zh-CN/docs/Web/JavaScript/Reference/Global_Objects/decodeURI)
- [decodeURIComponent()](https://developer.mozilla.org/zh-CN/docs/Web/JavaScript/Reference/Global_Objects/decodeURIComponent)
- [encodeURI()](https://developer.mozilla.org/zh-CN/docs/Web/JavaScript/Reference/Global_Objects/encodeURI)
- [encodeURIComponent()](https://developer.mozilla.org/zh-CN/docs/Web/JavaScript/Reference/Global_Objects/encodeURIComponent)
- [RegExp()](https://developer.mozilla.org/zh-CN/docs/Web/JavaScript/Reference/Global_Objects/RegExp)
- ...更多内置函数

## 全局对象： ​

- Application(DB OpenApi)

## 日志输出： ​

- console.log
- console.error
- console.info

## 自动补全： ​

支持自定义函数的自动补全功能。支持[JsDoc](https://jsdoc.app/about-getting-started)。

JavaScript

```JavaScript
/**
 * 数字增加1.
 *
 * @param {number} input 被加数.
 * @return +1.
 * @customfunction
 */
function inc(input) {
  return input + 1;
}
```

## 更多内置函数 ​

下面这些内置函数是用来帮助开发者处理字符串编码/解码、信息处理、参数获取和其他杂项任务的实用函数

## Crypto ​

对信息进行加密，摘要处理

### 示例 ​

javascript

```javascript
// 摘要foo这个字符串信息
console.log(Crypto.createHash("md5").update("foo").digest("hex")) // acbd18db4cc2f85cedef654fccc4a4d8
console.log(Crypto.createHmac("sha256", "a secret").update('some data to hash').digest('hex')) //7fd04df92f636fd450bc841c9418e5825c17f33ad9c87c518115a45971f7f77e
```

### 方法列表 ​

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| createHash(algorithm) | hash | 创建摘要算法实例，允许"md5", "sha1", "sha", "sha256", "sha512" |
| createHmac(algorithm, key) | hmac | 创建HMAC算法实例，允许"md5", "sha1", "sha", "sha256", "sha512" |

## hash ​

摘要对象，由Crypto产生

### 方法列表 ​

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| update(data[ ,inputEncoding]) | hash | 使用给定的 data 更新哈希内容，如果未提供 encoding，且 data 是字符串，则强制为 'utf8' 编码，如果 data 是 Buffer,则忽略 inputEncoding,可重复调用添加数据 |
| digest([encoding] | string| Buffer | 计算传给被哈希的所有数据的摘要，如果提供了 encoding，则将返回字符串；否则返回 Buffer。 |

## hmac ​

hmac对象，由Crypto产生

### 方法列表 ​

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| update(data[ ,inputEncoding]) | hash | 使用给定的 data 更新hmac内容，如果未提供 encoding，且 data 是字符串，则强制为 'utf8' 编码，如果 data 是 Buffer,则忽略 inputEncoding,可重复调用添加数据 |
| digest([encoding] | string| Buffer | 计算传给被hmac的所有数据的摘要，如果提供了 encoding，则将返回字符串；否则返回 Buffer。 |

## Buffer ​

产生一个 Buffer 实例

### 示例 ​

javascript

```javascript
// 创建包含字符串 'buffer' 的 UTF-8 字节的新缓冲区。
const buf = Buffer.from([0x62, 0x75, 0x66, 0x66, 0x65, 0x72]); 
console.log(buf.toString()) // buffer
```

### 方法列表 ​

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| from(array) | Buffer | 使用 0 – 255 范围内的字节 array 分配新的 Buffer。 |
| from(string[, encoding]) | Buffer | 从字符串转化为Buffer |
| from(arrayBuffer[, byteOffset[, length]]) | Buffer | 截断arrayBuffer的部分字节，生成新的Buffer |

## Time ​

时间函数，提供如休眠的方法

### 示例 ​

javascript

```javascript
Time.sleep(1000) // 休眠一秒
```

### 方法列表 ​

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| sleep(millisecond) | undefined | 休眠指定毫秒数 |

## Arguments ​

方便获取配置的参数数据

### 示例 ​

javascript

```javascript
Arguments.get("foo.bar", "defaults") // 如果自定义参数是{foo : {bar : "value"}}，则返回"value"，如果不存在，则返回第二个参数"defaults"
```

### 方法列表 ​

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| get(string[, defaults]) | any | 通过获取自定义参数的值，key支持使用.进行多次查找，如a.b会寻找{a : {b : ""}}这个结构体的b值。可指定默认值，如果找不到key对应的自定义参数，就返回默认值，没有指定默认值也找不到key返回undefined |
 
# 048 在线脚本AirScript / 脚本经典案例

本页内容

# 脚本经典案例 ​

## 例子1：选中区域快速批量填值 ​

javascript

```javascript
function main(){
  const time = getNowTime()
  const date = getNowDate()
  ActiveView.Selection(null, ["@日期", "@时间", "@分类"]).Value = [date, time, 'B']
}
// 获取当前时间，格式为 "hh:mm:ss"
function getNowTime() {
  return (new Date()).toTimeString().split(" ")[0]
}
// 获取当前日期，格式为 "yyyy:MM:dd"
function getNowDate() {
  const date = new Date()
  return date.getFullYear() + '/' + (date.getMonth() + 1) + '/' + date.getDate()
}

main()
```

## 例子2：快速实现“一键归档” ​

下面代码实现了一个文件中两张数据结构相同的表 把表一中的已完成的数据插入到表二中，并删除表一中数据 表结构如下图所示：

![例子2](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/db-airscript2.DTlAnBUb.png)

javascript

```javascript
function main() {
  const criterias = []
  criterias.push(Criteria("@分类",  "Equals", ["B"]))
  criterias.push(Criteria("@完成",  "Equals", ["1"]))
  // 创建filters
  const filters = []
  const filter = {Criterias: criterias, Op: "AND"}
  filters.push(filter)
  const range1 = Sheets(1).Views(1).RecordRange.Condition(filters, "AND")
  if(!range1){
    return
  }
  const length = range1.Count
  const range2 = Sheets(2).Views(1).RecordRange.Add(1, undefined, length)
  range2.Value = range1.Value
  range1.Delete()
}

main()
```

> 
> 结合上面两个例子，可以实现自动设置归档日期和时间，或者选中记录一键归档等等功能

## 例子3：快速删除空白数据 ​

删除名称字段中 值为空的数据

javascript

```javascript
function deleteRecords() {
  const criterias = []
  criterias.push(Criteria("@名称",  "Equals",['']))
  const filters = []
  const filter = {Criterias: criterias, Op: "AND"}
  filters.push(filter)
  Sheets(3).Views(1).RecordRange.Condition(filters, "AND").Delete()
}
deleteRecords()
```

## 例子4：快速创建一张表 ​

javascript

```javascript
function main() {
  Application.Sheets.Add(
    {
        Type:'xlEtDataBaseSheet',
        Config:{
            fields:
                [
                    {fieldType:'SingleLineText',args:{fieldName:'文本',fieldWidth:15}},
                    {fieldType:'MultiLineText',args:{fieldName:'多行文本',fieldWidth:15}},
                    {fieldType:'Date',args:{fieldName:'日期',numberFormat:'yyyy/mm/dd;@',fieldWidth:15}},
                    {fieldType:'SingleSelect',args:{fieldName:'单选项',fieldWidth:15,listItems:[{value: '选项1', color: 4283466178},{value: '选项2',color: 4281378020}]}},
                    {fieldType:'Number',args:{fieldName:'数字',fieldWidth:15}},
                    {fieldType:'Rating',args:{fieldName:'等级',maxRating:6,fieldWidth:15}},
                ],
            name:'数据表',
            views:
                [
                    {name:'表格视图',type:'Grid'},
                    {name:'表单视图',type:'Form'}
                ]
        }
    })
}
main()
```

## 例子5：格式化数据批量插入 ​

javascript

```javascript
function main(){
    const range = Application.Sheets(4).Views(1).RecordRange.Add(1, undefined, 300)// 在第1行，向上方添加300条记录
    const template = ["商品", 10]
    const range1 = []
    // 给1-300行赋值
    for (let i = 1; i < 301; i++ ) {
        if(i<101){
          range1.push([template[0]+i,template[1],'A'])
        }else if(i<201){
          range1.push([template[0]+i,template[1]+10,'B'])
        }else{
          range1.push([template[0]+i,template[1]+10,'C'])
        }
    }
    range.Value = range1
}
main()
```

## 例子 6：自动双向关联 ​

先来看一下表结构

- 客户表

![例子6](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/db-airscript6.Cz0SOEjV.png)

- 拜访记录表

![例子6](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/db-airscript6-2.Dbs6T-o0.png)

对 【客户表】 中的"拜访记录" 和 【拜访记录表】 中的 "客户详情"做自动关联

> 
> 通过客户名称来进行匹配

javascript

```javascript
function main() {
  const clientsView = Application.Sheets(1).Views(1)
  const clientsCount = clientsView.RecordRange.Count
  
  const visitsView = Application.Sheets(2).Views(1)
  const visitsCount = visitsView.RecordRange.Count

  const clients = clientsView.RecordRange('1:'+clientsCount)
  const visits = visitsView.RecordRange('1:'+visitsCount)

  linkVisits(clients, visits)
  linkClients(clients, visits)
}
// 关联拜访记录表中的客户
function linkClients(clients,visits){
  let name1 = ''
  let name2 = ''
  let id = 0
  // {"value":[{"id":"XP","str":"金山办公"}]}
  for(let i=1;i<visits.Count+1;i++){
    name1 = visits.Item(i,['@客户名称']).Value
    if(name1!==''){
      for(let j=1;j<clients.Count+1;j++){
            name2 = clients.Item(j,['@客户名称']).Value
            id = clients.Item(j).Id
            if(name1 === name2){
              // console.log({id:`${id}`,str:`${name2}`})
              visits.Item(j,['@客户详情']).Value = Application.DBCellValue([id])
              break
            }
      }
    }
    
  }
}
// 关联客户表中的拜访记录
function linkVisits(clients,visits){
  let name1 = ''
  let name2 = ''
  let id = 0
  
  // {"value":[{"id":"XP","str":"金山办公"}]}
  for(let i=1;i<clients.Count+1;i++){
    let vits = []
    name1 = clients.Item(i,['@客户名称']).Value
    if(name1!==''){
      for(let j=1;j<visits.Count+1;j++){
        name2 = visits.Item(j,['@客户名称']).Value
        id = visits.Item(j).Id
        if(name1 === name2){
          vits.push({id:`${id}`,str:`${name2}`})
          // visits.Item(i,['@客户详情']).Value = Application.DBCellValue([{id:`${id}`,str:`${name2}`}])
        }
      }
    }
    // console.log(vits)
    clients.Item(i,['@拜访记录']).Value = Application.DBCellValue(vits)
  }
}

main()
```

## 例子7：同步主表数据到其他表 ​

表结构如下

![例子7](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/db-airscript7.JtoBi7zl.png)

> 
> 第一张表中存着所有数据，其他几张表中存着分类数据 当主表中数据更新的时候，期望其他表数据也更新

javascript

```javascript
// 读取整个表
function getAllRecords(sheetIndex) {
  const view = Application.Sheets(sheetIndex).Views(1)
  const count = view.RecordRange.Count
  return view.RecordRange("1:"+count)
}
// 获取子表中匹配的同步字段数组
function getMatchFields(sheetIndex,toSyncFields){
  const matchFields = []
  const fieldDescs = Application.Sheets(sheetIndex).FieldDescriptors
  for(let i=0;i<toSyncFields.length;i++){
    const fieldName = toSyncFields[i]
    for(let i=1;i<fieldDescs.Count+1;i++){
      if(fieldDescs.Item(i).Name === fieldName){
        matchFields.push('@'+fieldName)
        break
      }
    }
  }
  return matchFields
}
// 检查记录值相同与否, 返回需要更新的记录
function checkFields(mainRecord, record, params) {
  const v1 = mainRecord.Item(1,params).Value
  const v2 = record.Item(1,params).Value
  if(v1===v2) return null
  // console.log(record.Item(1,params).Id,v1)
  return {id:record.Item(1,params).Id[0],val: params.length>1?v1:[v1]}
}

// 同步副表数据
function syncSheet(sheetIndex, mainRecordMap, {
  keyField,
  toSyncFields,
}) {
  const records = getAllRecords(sheetIndex)
  const matchFields = getMatchFields(sheetIndex,toSyncFields)
  // console.log(matchFields)
  const toUpdateIds = []
  const toUpdateVals = []
  for (let i = 1; i < records.Count+1; i++) {
    const record = records.Item(i)
    const name = records.Item(i,['@'+keyField]).Value
    if(name === '') continue
    const mainRecord = mainRecordMap[name]
    if (mainRecord) {
      const updatedResult = checkFields(mainRecord, record, matchFields)
      if (updatedResult) {
        toUpdateIds.push(updatedResult.id)
        toUpdateVals.push(updatedResult.val)
      }
    } else {
      console.error('没有在主表里面找到此条记录： ', name)
    }
  }
  if (toUpdateIds.length > 0) {
    // console.log(toUpdateVals)
    records(toUpdateIds,matchFields).Value = toUpdateVals
  }
}

// 同步主表数据到其他表
function syncMainSheetToOthers(mainSheetName, keyField, toSyncFields) {
  const sheets = Application.Sheet.GetSheets()
  const mainIndex = sheets.findIndex(item => item.name === mainSheetName)+1
  const mainRecords = getAllRecords(mainIndex)
  const recordsMap = {}
  for (let i = 1; i < mainRecords.Count+1; i++) {
    const name = mainRecords.Item(i,'@'+keyField).Value
    if(name === '') continue
    if (recordsMap[name]) {
      console.error('有重复的记录', name)
    }
    recordsMap[name] = mainRecords.Item(i)
  }
  for(let index=1;index<sheets.length+1;index++){
    if(index === mainIndex) continue
    syncSheet(index, recordsMap, {
      keyField,
      toSyncFields,
    })
  }
}

syncMainSheetToOthers('关爱清单', '产品名', ['机制', '价格'])
```

## 例子8：获取日期筛选后记录 ​

表结构如下

![例子8](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/db-airscript8.BbR_mp-w.png)

使用日期筛选进行获取日期为本月（2023年2月）的记录，即将筛选参数改为thisMonth，如"dynamicType": "thisMonth",筛选得到的记录可查看all变量

javascript

```javascript
function main(){
  const criterias = []
  criterias.push(Criteria("@日期",  "Equals", [{"dynamicType": "thisMonth","type": "DynamicSimple"}]))
  // 创建filters
  const filters = []
  const filter = {Criterias: criterias, Op: "AND"}
  filters.push(filter)
  const range1 = Application.Sheets(5).Views(1).RecordRange.Condition(filters, "AND")
  console.log(range1.Value)
  return range1
}
main()
```

## 例子9：获取联系人筛选后记录 ​

表结构如下

![例子9](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/db-airscript9.QCrgggWB.png)

使用联系人筛选进行获取联系人为Lin的记录，即更改筛选参数value改为“Lin”，如 "values": [ "Lin" ]，注意value值为数组形式,筛选得到的记录可查看all变量

javascript

```javascript
function filterContact() {
    const criterias = [Criteria("@联系人", "Equals", ["WPS_1719228187"])]
    // 创建filters
    const filters = [{Criterias: criterias, Op: "AND"}]
    const all = Application.Sheets(7).Views(1).RecordRange.Condition(filters)
    console.log(all.Value)
    return all
}

filterContact()
```

## 例子10：批量将获取的图片url进行文字识别 ​

![例子10](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/db-airscript10.DXb29cHV.png)

- 可以通过获取的图片url来进行调用百度云文字识别接口，识别图片中的文字

获取所有记录 ，获取该表所有图片URL，调用百度云文字识别接口，识别图片中的文字，并将结果生成在相应记录的“识别文字”字段下。

javascript

```javascript
const access_token="xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"

function main() {
    const view = Application.ActiveSheet.Views(1)
    const count = view.RecordRange.Count
    // 获取所有记录的图片附件值
    const attachments =  view.RecordRange('1:'+count, ["@图片和附件"]).Value
    // console.log(attachments)
    // 获取对应的url
    const urls = []
    let attachment = []
    // 这里的attachments是一个二维数组
    for(let i=0;i<attachments.length;i++){
        urls[i] = []
        attachment = attachments[i][0].Value
        for(let j = 0;j<attachment.length;j++){
          urls[i].push(attachment[j].LinkUrl || null)
        }
    }
    // console.log(urls)
    let words = []
    let url = []
    for(let i=0;i<urls.length;i++){
      url = urls[i]
      if(url){
        words[i] = [getText(url)]
      }else{
        words[i] = ['']
      }
      // console.log(words[i])
    }
    // console.log(recordIds,words)
    // 赋值到对应的 ‘识别文字’ 这一栏
    view.RecordRange('1:'+count,["@识别文字"]).Value = words
}
// 调用接口把所有的url转换为文字
function getText(urls){
    const word = urls.reduce((word,url)=>{
      // console.log('调用1次')
        let resp = HTTP.fetch('https://aip.baidubce.com/rest/2.0/ocr/v1/general_basic?access_token='+access_token,{
                method:"POST",
                timeout:2000,
                headers:{
                  'Content-Type': 'application/x-www-form-urlencoded',
                  'Accept': 'application/json',
                  'Authorization': access_token
            },
                body: `url=${encodeURIComponent(url)}`
            })
        if(resp.status !== 200){
           throw new Error("fetch err! status is "+resp.status)
        }
        return word + resp.text()
    },'')
    return word
}
main()
```

# 脚本示例（08/26更新） ​

javascript

```javascript
function main(){
  addSheetViewField()
  setDemoValue()
  testRecordRange()
  testFilters()
  testSorts()
  testGroups()
  testGridView()
}
function addSheetViewField(){
    const sheet = Application.Sheets.Add({After:1,Type:"xlEtDataBaseSheet"})
    sheet.Views.Add("Grid",'新建的表格视图')
    addField("MultiLineText","富文本")
    addField("Time","时间")
    addField("Currency","货币")
    addField("Percentage","百分比")
    addField("ID","身份证")
    addField("Phone","电话")
    addField("Email","邮箱")
}
function setDemoValue(){
  setValue("@文本","100")
  setValue("@文本","200",'4:6')
  setValue(2,123)
  setValue(2,456,'4:6')
  setValue(3,"2024/07/13")
  setValue(3,"2024/08/13","4:6")
  setValue("@单选项","选项1")
  setValue("@单选项","选项2","4:6")
  setValue(6,3)
  setValue(6,4,"4:6")
  setValue("@富文本","dfsafdasfdasf")
  setValue("@富文本","fdsfagzgzgfgvcbc","4:6")
  setValue("@时间","8:48:16")
  setValue("@时间","10:48:16","4:6")
  setValue("@货币","123213")
  setValue("@货币","1434434","4:6")
  // setValue(10,1.23)
  // setValue(10,2.35,"4:6")
  setValue("@身份证","43211620050917691X")
  setValue("@身份证","43211620050917691X","4:6")
  setValue("@电话","18064038091")
  setValue("@电话","18064838691","4:6")
  setValue("@邮箱","1437252712@qq.com")
  setValue("@邮箱","1437257777@qq.com","4:6")
}
function addField(type,name,index,valueUnique,defaultValue,defaultValueType,numberFormat,arg1,arg2,arg3,arg4){
    const desc = Application.Sheets(2).FieldDescriptors.FieldDescriptor(type, name)
    if (type === 'SingleSelect' || type === 'MultipleSelect') {
        desc.Items = arg1
    }
    if(type === 'Rating'){
        desc.MaxRating = arg1
    }
    if(type === 'Formula'){
        if (arg1) {
          desc.ValueUnique = arg1
         }if (arg2) {
             desc.ValueType =arg2
         }if (arg3) {
             desc.ShowPercentAsProgress = arg3
         }
    }
    if(type === 'Cascade'){
        const options = Application.CascadeOptions()
         if(typeof arg1 ==='object'){
            const o1 =  options.Add(arg1[0])
            const children = arg1[1]
            for(let i=0;i<children.length;i++){
              o1.Children.Add(children[i])
            }
          }
    if(typeof arg2 ==='object'){
        const o2 =  options.Add(arg2[0])
        const children = arg2[1]
        for(let i=0;i<children.length;i++){
           o2.Children.Add(children[i])
        }
    }
        desc.AllCascadeOption = options
    }
    if (type === 'OneWayLink') {
        desc.LinkSheet = arg1
        desc.IsAutoLink = arg2
    }
    if (valueUnique) {
        desc.ValueUnique = valueUnique
    }
    if (defaultValue) {
        desc.DefaultValue = defaultValue
    }
    if (defaultValueType) {
        desc.DefaultValue = defaultValueType
    }
    if (numberFormat) {
        desc.NumberFormat = numberFormat
    }
    const result =  Application.Sheets(2).FieldDescriptors.AddField(desc, index)
    if(result.Code !== 0){
      console.error(result.Message)
    }
}
function setValue(field,value,index="1:3"){
   if(typeof value === 'object'){
      value = Application.DBCellValue(value)
   }
   Application.Sheets(2).Views(1).RecordRange(index, field).Value = value
   const newVal = Application.Sheets(2).Views(1).RecordRange(index, field).Value
   let setSuccess = true
   if(typeof newVal === 'object'){
     for(let i=0;i<newVal.length;i++){
       if(newVal[i][0] === value){
          setConfirm = false
          break
       }
     }
   }else if(newVal !== value){
     setConfirm = false
   }
   if(!setSuccess)   console.error('setValue异常！')
}
function testRecordRange(){
  // RecordRange: Add、Item、Condition、Delete
  const rr = Application.Sheets(2).Views(1).RecordRange.Add(1,undefined,2)
  rr.Item(1).Value = ["300",789,"2024/09/13","选项1",[],3,"fadfadsf","10:33:33"]
  rr.Item(2).Value = ["400",789,"2024/09/13","选项1",[],3,"fadfadsf","10:33:33"]
  const criterias = [Criteria("@文本", "Equals",['400'])]
  const filters = [{Criterias: criterias, Op: "AND"}]
  const recordRange = Application.Sheets(2).Views(1).RecordRange.Condition(filters)
  if(recordRange){
    recordRange.Delete()
  }
}
function testFilters(){
  //Filters: Add、Item、Clear
  //Filter:  Delete
  Application.Sheets(2).Views(1).Filters.Clear()
  const filters = Application.Sheets(2).Views(1).Filters;
  const criteria = Criteria({
      Field: 1,
      CriteriaOp: 'Equals',
      Values: ["200"]
  })
  const filter = filters.Add(criteria);
  const delResult = filter.Delete()
  if(delResult.Code !== 0) {
    console.error(delResult.Message)
  } 
  const criteria1 = Criteria({
      Field: 1,
      CriteriaOp: 'Equals',
      Values: ["300"]
  })
  Application.Sheets(2).Views(1).Filters.Add(criteria1)
  const clsResult = Application.Sheets(2).Views(1).Filters.Clear()
  if(clsResult.Code !== 0) {
    console.error(clsResult.Message)
  } 
  const criteria2 = Criteria({
      Field: 1,
      CriteriaOp: 'Equals',
      Values: ["100"]
  })
  Application.Sheets(2).Views(1).Filters.Add(criteria2)
  //console.log(Application.Sheets(2).Views(1).Filters.Item(1).Criteria)
}
function testSorts(){
  //Sorts: Add Item Move 
  //Sort:  Delete ChangeField
  const sort= Application.Sheets(2).Views(1).Sorts.Add(2,true)
  if(sort.IsAscending !== true) {
    console.error('Add Sort时 设置IsAscending 异常')
  }
  Application.Sheets(2).Views(2).Sorts.Add(3,false)
  const changeRes = Application.Sheets(2).Views(1).Sorts.Item(1).ChangeField(1)
  if(changeRes.Code !== 0) {
    console.error(changeRes.Message)
  }
  Application.Sheets(2).Views(1).Sorts.Move([3,1])
  Application.Sheets(2).Views(1).Sorts.Add(2,false)
  const delRes = Application.Sheets(2).Views(1).Sorts.Item(2).Delete()
  if(delRes.Code !== 0) {
    console.error(delRes.Message)
  }
}
function testGroups(){
  // Groups: Add FoldAll UnFoldAll Item StatisticResult
  // Group:  ChangeField Delete
  const group = Application.Sheets(2).Views(1).Groups.Add(1,true)
  if(group.IsAscending !== true) {
    console.error('groups add 方法设置IsAscending 异常')
  }
  Application.Sheets(2).Views(1).Groups.Add(2,false)
  Application.Sheets(2).Views(1).Groups.FoldAll()
  const changeResult = Application.Sheets(2).Views(1).Groups.Item(2).ChangeField(3)
  if(changeResult.Code !== 0) {
    console.error(changeResult.Message)
  } 
  Application.Sheets(2).Views(1).Groups.Add(2)
  const delResult = Application.Sheets(2).Views(1).Groups.Item(2).Delete()
  if(delResult.Code !== 0) {
    console.error(delResult.Message)
  } 
  // console.log(Application.Sheets(2).Views(1).Groups.StatisticResult(1))
  Application.Sheets(2).Views(1).Groups.UnFoldAll()
}
function testGridView(){
    const gridView = Application.Sheets(2).Views.Add('Grid', '表格视图');
    if(!gridView.RowHeight){
       console.error('Views.Add表格视图时，返回值异常')
    }
    const itemGridView = Application.Sheets(2).Views(1)
    if(!itemGridView.RowHeight){
        console.error('Views.Item获取表格视图时，返回值异常')
    }
    const activeGridView = Application.ActiveView
    if(!activeGridView.RowHeight){
        console.error('ActiveView获取表格视图时，返回值异常')
    }
    Application.ActiveView.RowHeight = 'Medium';
    if(Application.ActiveView.RowHeight !== 'Medium'){
        console.error('设置GridView的RowHeight结果时异常')
    }
}
main()
```

> 
> 更多API用法可参考[多维表API](/documents/app-integration-dev/guide/dbsheet/Api/api-instro.html)
 
# 049 内嵌使用SDK / 简介

本页内容

# SDK 简介 ​

SDK提供了丰富的 **API** 对各类文档进行操作，通过使用 SDK，网页开发者可以自定义文档界面的元素、操作文档的内容、监听文档事件等操作，SDK 为用户提供了优质的在线文档体验。

接入方需要通过引入 SDK， 生成 WebOffice 文档的 `iframe`元素，将在线文档页面在宿主网页中展示出来。接入方网页和 WebOffice 文档、SDK 之间的通信过程如下：

![alt text](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/weboffice-instro1.D-iC7tx7.png)

## SDK 能做什么​ ​

开发者在引入并成功初始化 SDK 后，便可以使用 `SDK` 实例来帮助您完成文档相关的需求了。实例主要包含四个方面的能力：

- 在您的业务网页内显示 WebOffice 文档
- 通过在初始化时（调用WebOfficeSDK.init()函数时）灵活传递初始化配置，可以自定义文档界面、获取页面状态等
- 通过ApiEvent对象对文档的各类事件进行监听
- 通过Application对象对文档进行丰富的 API 调用，直接对文档的内容、格式和图形等进行操作变换

## SDK 1.0 版本和3.0 版本 有什么区别 ​

- JSSDK 1.0 通过 **iframe** 技术加载在线文档，有着更广泛的应用场景和丰富的功能适配
- JSSDK 3.0 版本通过 **微前端** 技术加载在线文档，有着跟客户系统一体化的体验

## 集成效果 ​

### 1. SDK 1.0 版本 ​

在智能文档场景下，借助 SDK 1.0 版本嵌入 WPS多维表格 的效果展示如下：

![alt text](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/weboffice-instro3.BGqUxb80.png)

#### 特点： ​

- 直接加载已有的云文档链接，WPS多维表格的交互被限定于既定视图区域内，无法超出区域。
- 通常情况下，文档内部划定一大块区域，用于加载类知识库内容，以此实现与业务系统的融合。

### 2. SDK 3.0 版本 ​

在智能文档场景下，借助 SDK 3.0 版本嵌入 WPS多维表格 的效果展示如下：

![alt text](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/weboffice-instro2.C0JR00IA.png)

#### 特点： ​

- WPS多维表格以嵌入形态融入文档，与业务系统紧密贴合，为用户提供一体化的交互体验，操作更为流畅。

> 
> 具体接入可以查看[快速入门](/documents/app-integration-dev/guide/dbsheet/Weboffice/weboffice-quickstart.html)
 
# 050 内嵌使用SDK / 快速入门

本页内容

# 开始​ ​

## 下载Demo示例 ​

1. 我们为您提供了一个Demo示例，建议下载使用，下载地址：[Demo示例](https://365.kdocs.cn/view/l/cb11ZL018X1g)
2. 下载 Demo 示例文件后，运行 npm install && npm run dev 启动

> 
> 下面示例代码都已在 Demo 中对应示例，推荐对照 Demo 示例查看

### 1. 内嵌接入（ SDK 1.x ）版本 ​

- 本质是 **iframe** 嵌套一个WPS多维表格
- 在 url 后面拼接 `readonly` 参数设置文件只读状态

#### 基本使用 ​

TypeScript

```TypeScript
import { Component, createRef } from 'react'
import WebOfficeSDK_V1 from '../weboffice-sdk/web-office-sdk-v1.1.11.es'

export default class PreviewDB extends Component {
    instance: any
    container = createRef<HTMLDivElement>()

    componentDidMount(): void {
        console.error(this.container)
        this.instance = WebOfficeSDK_V1.config({
            officeType: 'd',
            url: 'https://www.kdocs.cn/l/cuI4w9PX4CwI?disablePlugins&simple=1&readonly',
            mount: this.container.current,
        })
    }

    render() {
        return <div ref={this.container} style={{ height: '400px', width: '800px' }}></div>
    }
}
```

- 这样就在 div 容器里面加载了一个多维表格文件
- 通过传入参数配置，控制加载内容的展示与屏蔽，以及各项功能的开启与关闭。更多详细配置可查看：[更多配置](/documents/app-integration-dev/guide/dbsheet/Weboffice/weboffice-config-api.html)

### 2. 内嵌接入（ SDK 3.x ）版本 ​

#### 环境准备 ​

1. 保持 主页面 与 嵌入子文档页面 处于同域状态
2. 出现不同域时，要求服务端支持 CORS（跨域资源共享）机制

> 
> 如：主文档使用 https://test.kdocs.cn， 嵌入子文档页面使用 docs.kdocs.cn docs.kdocs.cn 的服务端设置：Access-Control-Allow-Origin: https://test.kdocs.cn

其他头部配置可以参考 kdocs 配置：

javascript

```javascript
Access-Control-Allow-Origin: https://test.kdocs.cn
Access-Control-Allow-Credentials: true
Access-Control-Allow-Headers: accept, content-type, x-user-query, x-device-id, x-requested-with, x-csrftoken, accept-encoding, accept-language, x-csrf-rand, x-server-id, x-endpoint-id, x-app-id, EagleEye-TraceID, EagleEye-SessionID, EagleEye-pAppName, x-user-token, x-app-id, x-app-token, x-redirect-region-group
Access-Control-Allow-Methods: GET, POST, PUT, DELETE, OPTIONS
Access-Control-Expose-Headers: accept, content-type, x-user-query, x-device-id, x-requested-with, x-csrftoken, accept-encoding, accept-language, x-csrf-rand, x-server-id, x-endpoint-id, x-app-id, EagleEye-TraceID, EagleEye-SessionID, EagleEye-pAppName, x-user-token, x-app-id, x-app-token
```

#### 基本使用 ​

TypeScript

```TypeScript
import { Component, createRef } from 'react'
import WebOfficeSDK_V3 from '../weboffice-sdk/web-office-sdk-v3.4.1.es'

export default class EmbedDB extends Component {
    instance: any
    container = createRef<HTMLDivElement>()

    componentDidMount(): void {
        this.instance = WebOfficeSDK_V3.config({
            officeType: 'd',
            // embed=1 开启内嵌模式
            // disablePlugins 关闭插件
            url: 'https://365.kdocs.cn/l/cuI4w9PX4CwI?embed=1&disablePlugins',
            mount: this.container.current,
        })
    }

    render() {
        return <div ref={this.container} style={{ height: '400px', width: '800px' }}></div>
    }
}
```

> 
> 这样就能够看到页面成功加载了一个多维表格文件
 
# 051 内嵌使用SDK / 基础用法 / 配置参数

本页内容

# 配置参数 ​

在初始化文档应用时，支持传入一些配置项，来对文档进行一些基本设置。

WPS多维表格常用配置：

javascript

```javascript
WebOfficeSDK.config({
        // embed 参数控制开启 WPS多维表格 适配过的嵌入模式，样式可以参考 智能文档嵌入 WPS多维表格
        // disablePlugins 关闭插件, 公网适用
        url: url + '?disablePlugins&' + (fullScreen ? 'embed-fullscreen=1' : 'embed=1'),
        viewMode: 'Embed',
        mode: 'embed',
        fileType: 'd',
        mount: container,
        hideGuide: true,
        debug,
        commonOptions: {
            isEnableChangeDocumentTitle: false,
            isShowHeader: false,
            disableSafeCs: true,
        },
        dbOptions: {
            embed: {
                commandBarSlotElm: slotElement,
                enableChangingView: true,
                enableToolBar: true,
            },
            enableRecordCooperation: false,
            enableFavView: false,
            // 字段白名单
            fieldsWhiteList: [
                'Checkbox',
                'Complete',
                'Contact',
                'CreatedTime',
                'Currency',
                'Date',
                'Email',
                'MultiLineText',
                'MultipleSelect',
                'Number',
                'Percentage',
                'Phone',
                'Rating',
                'SingleLineText',
                'SingleSelect',
                'Time',
                'Url',
                'Attachment',
                'ID'
            ],
            // 视图菜单白名单
            viewMenuFeatureWhiteList: [
                'rename',
                'addDesc',
                'deleteView',
                'copyView'
            ],
            // 附件类型白名单
            attachmentSuffixWhiteList: [
                'pptx', 'ppt', 'pot', 'potx', 'pps', 'ppsx', 'dps', 'dpt', 'pptm', 'potm', 'ppsm', 'xls',
                'xlt', 'et', 'ett', 'xlsx', 'ksheet', 'xltx', 'xlsb', 'xlsm', 'xltm', 'dbt', 'csv', 'lrc',
                'c', 'cpp', 'h', 'asm', 'java', 'asp', 'bat', 'bas', 'prg', 'cmd', 'txt', 'log', 'xml',
                'htm', 'html', 'pdf', 'doc', 'dot', 'wps', 'wpt', 'docx', 'dotx', 'docm', 'dotm', 'rtf',
                'png', 'jpg', 'jpeg', 'bmp', 'gif', 'otl', 'kw', 'ofd', 'pom', 'pof', 'heic'
            ],
        }
    })
```

## 配置参数(**config**)列表 ​

| 配置参数 | 类型 | 说明 |
| --- | --- | --- |
| url | string | 在线文档预览编辑地址，如 `https://www.kdocs.cn/l/cuI4w9PX4CwI` |
| mount | HTMLElement | 挂载节点，详见 [挂载节点](/documents/app-integration-dev/guide/dbsheet/Weboffice/weboffice-config-mount.html) |
| mode | string | 显示模式，使用内嵌模式：`embed` |
| viewMode | string | 视图模式，使用内嵌模式：`Embed` |
| fileType | string | 文件类型，WPS多维表格使用 `d` |
| commonOptions | object | 组件通用配置，详见 commonOptions |
| dbOptions | object | WPS多维表格自定义配置，详见 dbOptions |

## 组件通用配置(**commonOptions**)列表 ​

| 配置参数 | 类型 | 说明 |
| --- | --- | --- |
| isEnableChangeDocumentTitle | boolean | 是否禁止修改页面标题 |
| isShowHeader | boolean | 是否显示头部区域（顶部工具栏） |
| disableSafeCs | boolean | 是否关闭嵌入子文档安全区域 |

## WPS多维表格自定义配置(**dbOptions**)列表 ​

| 配置参数 | 类型 | 说明 |
| --- | --- | --- |
| embed | object | 嵌入模式下的配置，详见 embed |
| enableRecordCooperation | boolean | 是否开启行协作功能 |
| enableFavView | boolean | 是否屏蔽收藏视图功能 |
| fieldsWhiteList | string[] | 字段白名单 |
| viewMenuFeatureWhiteList | string[] | 视图菜单白名单 |
| attachmentSuffixWhiteList | string[] | 附件类型白名单 |

## embed配置(**embed**)列表 ​

| 配置参数 | 类型 | 说明 |
| --- | --- | --- |
| commandBarSlotElm | HTMLElement | 用户自定义传入的工具栏组件 |
| enableChangingView | boolean | 是否开启切换视图 |
| enableToolBar | boolean | 是否开启工具栏 |
 
# 052 内嵌使用SDK / 基础用法 / 挂载节点

本页内容

- 示例代码：
- iframe 对象

# 挂载节点 ​

挂载节点是指文档应用插入页面时挂载的 HTML DOM 节点，如果没有指定挂载节点，则默认生成一个撑满全屏的 div 节点。

JSSDK 初始化时，会自动在挂载节点下面插入一个 **iframe** 元素，并在该 **iframe** 元素中渲染文档应用。

> 
> 请在 `load` 或 `DOMContentLoaded` 事件被触发后，确保挂载节点存在，再执行初始化操作。

> 
> 由于 **iframe** 限制，需要在初始化时给挂载节点指定具体宽高，否则可能会导致文档绘制异常。

## 示例代码： ​

html

```html
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no" />
  <meta http-equiv="X-UA-Compatible" content="ie=edge" />
  <title>WPS WebOffice Demo</title>
  <style>
    * { margin: 0; padding: 0; }
    html, body { width: 100%; height: 100%; overflow: hidden; }
    .custom-mount { width: 100%; height: 100%; }
  </style>
</head>
<body>
  <div class="custom-mount"></div>

  <script src="./web-office-sdk-v1.1.8.umd.js"></script>
  <script>
    window.onload = function() {
      const jssdk = WebOfficeSDK.config({
        url: '在线文档预览编辑地址', // 该地址需要对接方服务端提供，形如 https://wwo.wps.cn/office/p/xxx
        mount: document.querySelector('.custom-mount'),
      })
    }
  </script>
</body>
</html>
```

## iframe 对象 ​

如果需要对 **iframe** 对象做特殊处理，可以通过 JSSDK 实例化对象快速获取到 **iframe** 的 **DOM** 对象。

javascript

```javascript
const jssdk = WebOfficeSDK.config({
  mount: document.querySelector('.custom-mount'),
});
console.log(jssdk.iframe);
```
 
# 053 内嵌使用SDK / 基础用法 / 事件处理

本页内容

# 监听事件 ​

## 公共事件列表： ​

| 事件名 | 说明 |
| --- | --- |
| fileOpen | 文档打开 |
| error | 错误事件 |
| fileStatus | 文件保存状态 |
| fullscreenchange | 进入或退出全屏事件 |

### fileOpen ​

文档打开成功或者失败时的事件回调

注意：该事件需要在 jssdk.ready() 之前进行注册

javascript

```javascript
instance.ApiEvent.AddApiEventListener("fileOpen", (data) => {
  console.log("fileOpen: ", data);
});
```

成功响应：

javascript

```javascript
{
  duration: 812,
  fileInfo: {
    createTime: 1606461829,
    id: "94749723688",
    modifyTime: 1606461829,
    name: "example.doc",
    officeType: "s",
  },
  stageTime: 1614,
  success: true,
  time: 1614,
  ts: 1607858260164,
}
```

失败响应：

javascript

```javascript
{
  msg: "Fail",
  result: "Fail"
}
```

### error ​

错误发生时的事件回调

例如将 doc 文件改成 xls 文件等操作，会引发报错

javascript

```javascript
instance.ApiEvent.AddApiEventListener("error", (data) => {
  console.log("error: ", data);
});
```

返回参数：

javascript

```javascript
{
  reason: "Fail";
}
```

### fileStatus ​

文件保存的事件回调

javascript

```javascript
instance.ApiEvent.AddApiEventListener("fileStatus", (data) => {
  console.log("fileStatus: ", data);
});
```

返回参数：

javascript

```javascript
{
  status: 0, // 文档无更新
  status: 1, // 版本保存成功, 触发场景：手动保存、定时保存、关闭网页
  status: 2, // 暂不支持保存空文件, 触发场景：内核保存完后文件为空
  status: 3, // 空间已满
  status: 4, // 保存中请勿频繁操作，触发场景：服务端处理保存队列已满，正在排队
  status: 5, // 保存失败
  status: 6, // 文件更新保存中，触发场景：修改文档内容触发的保存
  status: 7, // 保存成功，触发场景：文档内容修改保存成功
}
```

### fullscreenchange ​

进入或者退出全屏时会执行回调

如果在 commonOptions 下配置了 isBrowserViewFullscreen 或者 isIframeViewFullscreen，此项监听会无效。

javascript

```javascript
instance.ApiEvent.AddApiEventListener("fullscreenchange", (data) => {
  console.log("fullscreenchange: ", data);
});
```

返回参数：

javascript

```javascript
{
  status: 0, // 退出全屏
  status: 1, // 进入全屏
}
```

## WPS多维表格事件列表： ​

| 事件名 | 说明 |
| --- | --- |
| ViewTypeChanged | 监听视图变化 |
| ViewDataUpdate | 监听数据更新 |
| SelectionChange | 监听选区变化 |
| ActiveDetailRecordChange | 详情页卡片记录变化 |

使用方法如下：

javascript

```javascript
instance.Application.Sub.ViewTypeChange = (data) => {
  console.log('ViewTypeChange: ', data)
}
```
 
# 054 内嵌使用SDK / 基础用法 / 高级用法

本页内容

# 高级用法 ​

更多高级用法可参考[多维表API](/documents/app-integration-dev/guide/dbsheet/Api/api-instro.html)
 
# 055 内嵌使用SDK / 界面定制案例

本页内容

# 界面定制案例 ​

## 1. 控制视图变化 ​

首先封装一个类用来与 WPS多维表格 通信

TypeScript

```TypeScript
export default class DBInstance {
    instance: SDKInstance
    // 与 WPS多维表格 通信 实例 ，可以调用 WPS多维表格 的对外提供的 API
    Application: DBApplication

    async init(container: HTMLElement, fileUrl: string): Promise<any> {
        const instance = WebOfficeSDK_V3.config(getSDKConfig(container, fileUrl))
        await instance.commonApiReady()
        this.instance = instance
        this.Application = instance.Application
    }
    // 设置显示或隐藏 Toolbar
    async setToolBar(visible: boolean) {
        this.Application.ActiveDBSheet.SetNeedToolBar(visible)
    }
    // 获取多维表大小
    async getGridSize() {
        const size = await this.Application.ActiveDBSheet.View.GetViewRect()
        return size
    }

    destroy() {
        return this.instance.destroy()
    }
}
```

在 React 代码里控制：使当前 `div` 容器大小适配多维表，同时管控 `Toolbar` 的显隐。代码如下

tsx

```tsx
export default class EmbedInteractiveDB extends Component<{}, State> {
    instance: DBInstance
    container = createRef<HTMLDivElement>()
    constructor(props: {} | Readonly<{}>) {
        super(props)
        this.state = {
            height: 400,
            width: 800,
        }
    }
    async componentDidMount() {
        this.instance = new DBInstance()
        await this.instance.init(this.container.current, TEST_FILE_URL)
        // 获取多维表大小
        const size = await this.instance.getGridSize()
        // 设置当前div大小为多维表的大小
        this.setState({
            width: size.Width,
            height: size.Height,
        })
    }
    componentWillUnmount(): void {
        this.instance?.destroy()
    }
    render() {
        const { height, width } = this.state
        return (
            <div
                ref={this.container}
                style={{ height: `${height}px`, width: `${width}px` }}
                onMouseEnter={() => this.instance?.setToolBar(true)}
                onMouseLeave={() => this.instance?.setToolBar(true)}
            />
        )
    }
}
```

## 2. 监听事件 ​

> 
> 通过 SDK 监听 WPS多维表格 内的事件，进而执行相关操作

在 React 代码里增加事件监听代码，每当视图或数据更新时，就能自动适配页面大小，保障用户始终获得最佳视觉体验。

TypeScript

```TypeScript
async componentDidMount() {
    this.instance = new DBInstance()
    await this.instance.init(this.container.current, TEST_FILE_URL)
    const size = await this.instance.getGridSize()
    this.setState({
        width: size.Width,
        height: size.Height,
    })
    // 监听视图变化
    this.instance.Application.Sub.ViewTypeChanged = this.onViewTypeChanged
    // 监听数据更新
    this.instance.Application.Sub.ViewDataUpdate = this.onViewDataUpdate
}

onViewDataUpdate = ({ Width, Height }: Size) => {
    this.setState({
        width: Width,
        height: Height,
    })
}

onViewTypeChanged = async ({ viewType, id }) => {
    // 获取视图尺寸大小
    const size = await this.instance.getGridSize()
    this.setState({
        width: size.Width,
        height: size.Height,
    })
}

componentWillUnmount(): void {
    if (this.instance) {
        this.instance.Application.Sub.ViewDataUpdate = null
        this.instance.Application.Sub.ViewTypeChanged = null
        this.instance.destroy()
    }
}
```

> 
> 具体 Demo 可以参考 Event Demo， 源码在 `EmbedEventDB.tsx` 中 还可以监听一些视图变化的其他事件，具体事件在 `Application.Sub` 的属性中。 详细代码可以参考 Demo 文件，更多事件处理可查看[事件处理](/documents/app-integration-dev/guide/dbsheet/Weboffice/weboffice-config-event.html)

## 3. 控制滚动 ​

> 
> 当外部视图宽度达到最大的时候，需借助 WPS多维表格的左右滚动查看数据

下面给出一个实现左右滚动的示例：在 React 代码中添加对 wheel 事件的监听，在事件处理逻辑里操控 WPS多维表格 实现左右滚动。

TypeScript

```TypeScript
const MAX_WIDTH = 1000
export default class EmbedScrollingDB extends Component<{}, State> {
    instance: DBInstance
    container = createRef<HTMLDivElement>()
    hScroll = 0
    maxScroll = 0

    constructor(props: {} | Readonly<{}>) {
        super(props)
        this.state = {
            height: 400,
            width: MAX_WIDTH,
        }
    }

    async componentDidMount() {
        this.instance = new DBInstance()
        await this.instance.init(this.container.current, TEST_FILE_URL)
        this.instance.setEmbedMaxWidth(MAX_WIDTH)
        const size = await this.instance.getGridSize()
        this.updateSize(size.Width, size.Height)

        this.instance.Application.Sub.ViewTypeChange = this.onViewTypeChange
        this.instance.Application.Sub.ViewDataUpdate = this.onViewDataUpdate
        this.instance.Application.Sub.SelectionChange = this.onDBSelectionChange
    }

    updateSize = (width: number, height: number) => {
        if (width > MAX_WIDTH) {
            this.maxScroll = width - MAX_WIDTH
        }
        this.setState({
            width: width > MAX_WIDTH ? MAX_WIDTH : width,
            height,
        })
    }

    onDBSelectionChange = async () => {
        // 获取视图滚动条的位置
        const pos = await this.instance?.Application.ActiveDBSheet.View.GetScrollBar()
        if (pos) {
            this.hScroll = pos.X
        }
    }

    onViewDataUpdate = ({ Width, Height }: Size) => {
        this.updateSize(Width, Height)
    }

    onViewTypeChange = async ({ viewType, id }) => {
        const size = await this.instance.getGridSize()
        this.updateSize(size.Width, size.Height)
    }

    onMouseEnter = () => {
        if (this.instance?.Application) {
            this.instance?.setToolBar(true)
        }
    }

    onMouseLeave = () => {
        if (this.instance?.Application) {
            this.instance?.setToolBar(false)
        }
    }

    onWheel = ({ deltaX, deltaY }: WheelEvent) => {
        if (!this.instance?.Application) {
            return 
        }
        if (Math.abs(deltaX) > Math.abs(deltaY)) {
            this.hScroll += deltaX
            this.hScroll = Math.max(0, this.hScroll)
            this.hScroll = Math.min(this.maxScroll, this.hScroll)
            // 设置滚动条位置
            this.instance?.Application.ActiveDBSheet.View.SetScrollPos({ X:  this.hScroll, Y: 0 })
        }
    }

    componentWillUnmount(): void {
        if (this.instance) {
            this.instance.Application.Sub.ViewDataUpdate = null
            this.instance.Application.Sub.ViewTypeChange = null
            this.instance.Application.Sub.SelectionChange = null
            this.instance.destroy()
        }
    }

    render() {
        const { height, width } = this.state
        return (
            <div
                ref={this.container}
                style={{ height: `${height}px`, width: `${width}px` }}
                onMouseEnter={this.onMouseEnter}
                onMouseLeave={this.onMouseLeave}
                onWheel={this.onWheel}
            />
        )
    }
}
```

## 4. 自定义工具栏元素 ​

> 
> 实现自定义多维表格工具栏元素，如下图增加工具栏右侧`按钮操作`，定制个性化需求。

![alt text](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/weboffice3.1_9rbrcE.png)

1. 先创建自定义组件 `SlotElement`，组件内部的交互增加按钮操作，点击按钮处理相应事件

TypeScript

```TypeScript
export default class SlotElement extends React.Component {
    onClickButton1 = () => {
        console.log('点击了 按钮1')
    }
    onClickButton2 = () => {
        console.log('点击了 按钮2')
    }
    render() {
        return (
            <>
                <button onClick={this.onClickButton1}>按钮1</button>
                <button onClick={this.onClickButton2}>按钮2</button>
            </>
        )
    }
}
```

1. 在 SDK `config` 方法中的 `dbOptions` 传入 `SlotElement`，即可完成配置

TypeScript

```TypeScript
const instance = WebOfficeSDK_V3.config({
    url: url,
    viewMode: 'Embed',
    mode: 'embed',
    fileType: 'd',
    mount: container,
    dbOptions: {
        embed: {
            commandBarSlotElm: slotElement, // 自定义元素
            enableChangingView: true,
            enableToolBar: true,
        },
    },
    //其他配置
})
```
 
# 056 API文档 / 简介

本页内容

- 怎么使用 API
    - 浏览器环境示例
    - 脚本编辑器示例
    - 相同点
    - 差异点

# API简介 ​

WPS多维表格为开发者提供了一套功能完备的聚合 API 体系，借助配套的 [SDK](/documents/app-integration-dev/guide/dbsheet/Weboffice/weboffice-instro.html)，开发者既能够在浏览器环境中灵活调用 API 开展项目开发，也能依托 [在线脚本AirScript](/documents/app-integration-dev/guide/dbsheet/AirScript/AirScript-instro.html) 运用 API 来编写脚本，拓展业务功能。 丰富多样的 API 接口，赋予了用户极大的自主开发空间，扩展多维表的功能，定制个性化功能。

## 怎么使用 API ​

> 
> 以快速创建一张表为例，下面分别是在 **浏览器环境** 和 **AirScript脚本编辑器** 中的使用示例

### 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  await app.Sheets.Add(
    null,1,'xlEtDataBaseSheet',
        {
            fields:
                [
                    {fieldType:'SingleLineText',args:{fieldName:'文本',fieldWidth:15}},
                    {fieldType:'MultiLineText',args:{fieldName:'多行文本',fieldWidth:15}},
                    {fieldType:'Date',args:{fieldName:'日期',numberFormat:'yyyy/mm/dd;@',fieldWidth:15}},
                    {fieldType:'SingleSelect',args:{fieldName:'单选项',fieldWidth:15,
                        listItems:[{value: '选项1', color: 4283466178},{value: '选项2',color: 4281378020}]}},
                    {fieldType:'Number',args:{fieldName:'数字',fieldWidth:15}},
                    {fieldType:'Rating',args:{fieldName:'等级',maxRating:6,fieldWidth:15}},
                ],
            name:'数据表',
            views:
                [
                    {name:'表格视图',type:'Grid'},
                    {name:'表单视图',type:'Form'}
                ]
        }
    )
}
```

### 脚本编辑器示例 ​

javascript

```javascript
function main() {
  Application.Sheets.Add(
    1,null,'xlEtDataBaseSheet',
        {
            fields:
                [
                    {fieldType:'SingleLineText',args:{fieldName:'文本',fieldWidth:15}},
                    {fieldType:'MultiLineText',args:{fieldName:'多行文本',fieldWidth:15}},
                    {fieldType:'Date',args:{fieldName:'日期',numberFormat:'yyyy/mm/dd;@',fieldWidth:15}},
                    {fieldType:'SingleSelect',args:{fieldName:'单选项',fieldWidth:15,
                        listItems:[{value: '选项1', color: 4283466178},{value: '选项2',color: 4281378020}]}},
                    {fieldType:'Number',args:{fieldName:'数字',fieldWidth:15}},
                    {fieldType:'Rating',args:{fieldName:'等级',maxRating:6,fieldWidth:15}},
                ],
            name:'数据表',
            views:
                [
                    {name:'表格视图',type:'Grid'},
                    {name:'表单视图',type:'Form'}
                ]
        }
    )
}
main()
```

### 相同点 ​

- 两者都是采用 JavaScript 语言编写
- 采用同一套 **API** 体系，接口一致

### 差异点 ​

- 浏览器环境中，使用 Async、Await 语法来获取数据和设置数据，脚本编辑器中，不支持使用 Async、Await、Promise 语法
- 脚本编辑器中，内置了一些[基本数据类型、对象和函数](/documents/app-integration-dev/guide/dbsheet/AirScript/AirScript-build-in.html)用来帮助开发者开发
 
# 057 API文档 / 数据表 / 新建数据表

本页内容

# 新建数据表 ​

JSSDK: v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

新建数据表到指定位置，Before 和 After 只需要提供一个，另一个填 null 即可

## 语法 ​

表达式.Add(Before, After,Type,Config)

表达式：Sheets

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Before | 否 | number/string | 插入到Before(索引从1开始/数据表名)对应sheet之前 |
| After | 否 | number/string | 插入到After(索引从1开始/数据表名)对应sheet之后 |
| Type | 是 | string | "xlEtFlexPaperSheet"(说明页面)(暂不支持)、"xlEtDataBaseSheet"（数据表）、"xlDbDashBoardSheet"（仪表盘） |
| Config | 否 | object | 数据表专属配置，结构：Config:{ fields : Field[] , name ?: string , views ?: View[] }； |

## 参数Config属性详解 ​

| 属性名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| fields | 否 | Field[] | 字段数组，Field类型结构：{fieldType: FieldType,args: { fieldName: string, fieldWidth: number, listItems?: { value: string, color: number}[], numberFormat?: string, maxRating?: number } } |
| name | 否 | string | 数据表名,默认为‘data1’ |
| views | 否 | View[] | 视图配置数组，View结构：{name: string,type: ViewType}，ViewType的取值为：'Grid'（网格视图）、'Kanban'（看板视图）、'Gallery'（相册视图）、'Form'（表单视图）、'Gantt'（甘特视图）、'Query'（查询视图）或'Calendar'（日历视图）；默认创建'Grid'。暂只支持'Grid'和'Form'。 |

## 返回值 ​

Sheet

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  await app.Sheets.Add(
    null,1,'xlEtDataBaseSheet',
        {
            fields:
                [
                    {fieldType:'SingleLineText',args:{fieldName:'文本',fieldWidth:15}},
                    {fieldType:'MultiLineText',args:{fieldName:'多行文本',fieldWidth:15}},
                    {fieldType:'Date',args:{fieldName:'日期',numberFormat:'yyyy/mm/dd;@',fieldWidth:15}},
                    {fieldType:'SingleSelect',args:{fieldName:'单选项',fieldWidth:15,
                        listItems:[{value: '选项1', color: 4283466178},{value: '选项2',color: 4281378020}]}},
                    {fieldType:'Number',args:{fieldName:'数字',fieldWidth:15}},
                    {fieldType:'Rating',args:{fieldName:'等级',maxRating:6,fieldWidth:15}},
                ],
            name:'数据表',
            views:
                [
                    {name:'表格视图',type:'Grid'},
                    {name:'表单视图',type:'Form'}
                ]
        }
    )
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
  Application.Sheets.Add(
    1,null,'xlEtDataBaseSheet',
        {
            fields:
                [
                    {fieldType:'SingleLineText',args:{fieldName:'文本',fieldWidth:15}},
                    {fieldType:'MultiLineText',args:{fieldName:'多行文本',fieldWidth:15}},
                    {fieldType:'Date',args:{fieldName:'日期',numberFormat:'yyyy/mm/dd;@',fieldWidth:15}},
                    {fieldType:'SingleSelect',args:{fieldName:'单选项',fieldWidth:15,
                        listItems:[{value: '选项1', color: 4283466178},{value: '选项2',color: 4281378020}]}},
                    {fieldType:'Number',args:{fieldName:'数字',fieldWidth:15}},
                    {fieldType:'Rating',args:{fieldName:'等级',maxRating:6,fieldWidth:15}},
                ],
            name:'数据表',
            views:
                [
                    {name:'表格视图',type:'Grid'},
                    {name:'表单视图',type:'Form'}
                ]
        }
    )
}
main()
```
 
# 058 API文档 / 数据表 / 删除数据表

本页内容

# 删除数据表 ​

## 说明 ​

通过索引位置或数据表名来删除指定表

## 语法 ​

表达式.Delete(Index)

表达式: Sheets

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Index | 是 | number/string | 索引从 1 开始/数据表名 |

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    await app.Sheets.Delete(1);
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  Application.Sheets.Delete(1);
}
main()
```
 
# 059 API文档 / 数据表 / 移动数据表

本页内容

# 移动数据表 ​

## 说明 ​

移动数据表到指定位置，Before 和 After 只需要提供一个，另一个填 null 即可

## 语法 ​

表达式.Move(From, Before, After)

表达式: Sheets

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| From | 是 | number/string | 待移动 sheet 的名称或索引号，从 1 开始 |
| Before | 是 | number/string | 移动到 Before(索引从 1 开始/数据表名)对应 sheet 之前 |
| After | 是 | number/string | 移动到 After(索引从 1 开始/数据表名)对应 sheet 之后 |

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    await app.Sheets.Move(111111, null, 22222);
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  Application.Sheets.Move(111111, null, 22222);
}
main()
```
 
# 060 API文档 / 数据表 / 添加说明

本页内容

# 添加说明 ​

## 说明 ​

为当前数据表添加说明

## 语法 ​

表达式.AddDescription(Value)

表达式: Sheet

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Value | 是 | string | 待添加的说明文案 |

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    await app.Sheets(1).AddDescription('hello');
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  Application.Sheets(1).AddDescription('hello');
}
main()
```
 
# 061 API文档 / 数据表 / 设置图标

本页内容

# 设置图标 ​

## 说明 ​

可读写

设置图标、返回当前数据表的图标

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const sheet = app.Sheets(1);
    // read
    const sheetIcon = await sheet.Icon;
    // write
    sheet.Icon = '📚';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const sheet = Application.Sheets(1);
    // read
    const sheetIcon =  sheet.Icon;
    // write
    sheet.Icon = '📚';
}
main()
```
 
# 062 API文档 / 数据表 / 重命名

本页内容

# 重命名 ​

## 说明 ​

可读写

重命名、返回当前数据表的名称

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const sheet = app.Sheets(1);
    // read
    const sheetName = await sheet.Name;
    // write
    sheet.Name = 'newSheetName';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const sheet = Application.Sheets(1);
    // read
    const sheetName = sheet.Name;
    // write
    sheet.Name = 'newSheetName';
}
main()
```
 
# 063 API文档 / 数据表 / 创建副本

本页内容

# 创建副本 ​

## 说明 ​

为当前数据表创建副本

## 语法 ​

表达式.Copy(Value)

表达式: Sheet

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Value | 否 | boolean | 创建副本的方式，默认为 false。传 true 复制全部内容；传 false 仅复制空表和视图；复制的sheet为仪表盘时，此参数不传 |

## 返回值 ​

Sheet

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    // 复制全部内容
    await app.Sheets(1).Copy(true);
    // 仅复制空表和视图
    await app.Sheets(1).Copy(false);
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    // 复制全部内容
    Application.Sheets(1).Copy(true);
    // 仅复制空表和视图
    Application.Sheets(1).Copy(false);
}
main()
```
 
# 064 API文档 / 数据表 / 监听增加数据表的事件

本页内容

- 说明
- 语法
- 参数
- 返回值
- 事件返回数据
- 浏览器环境示例
- 脚本编辑器 示例

# 监听增加数据表的事件 ​

## 说明 ​

为当前数据表集合添加 CreateSheet 事件(当前只支持添加数据表事件，添加说明页和仪表盘不会触发该事件，后续版本更新后支持),当新增 sheet 时触发。注意在脚本编辑器中使用时，脚本运行结束就会退出运行，这时可能回调无法被正常触发

## 语法 ​

表达式: OnCreateSheet(Callback)

表达式: Sheets

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Callback | 是 | func | 指定事件发生时的回调方法 ; const eventContext = await Sheets.OnCreateSheet(()=&gt;{ ... }) |

## 返回值 ​

EventContext

## 事件返回数据 ​

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| Sheet | Sheet | 表 |

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    let eventContext;
    eventContext = await app.Sheets.OnCreateSheet(data => {
        console.log(data);
        // 取消事件监听
        eventContext.Destroy();
    });
    await app.Sheets.Add({ Type: 'xlEtDataBaseSheet' });
    //这里会执行OnCreateSheet的回调
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    let eventContext;
    eventContext = Application.Sheets.OnCreateSheet(data => {
        console.log(data);
        // 取消事件监听
        eventContext.Destroy();
    });
    Application.Sheets.Add({ Type: 'xlEtDataBaseSheet' });
    //这里会执行OnCreateSheet的回调
}
main();
```
 
# 065 API文档 / 数据表 / 监听重命名数据表的事件

本页内容

# 监听重命名数据表的事件 ​

## 说明 ​

为当前数据表添加 Rename 事件,当被修改 Name 时触发。注意在脚本编辑器中使用时，脚本运行结束就会退出运行，这时可能回调无法被正常触发

## 语法 ​

表达式.OnRename(Callback)

表达式: Sheet

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Callback | 是 | func | 指定事件发生时的回调方法 ; const eventContext = await Sheet.OnRename(()=&gt;{ ... }) |

## 返回值 ​

EventContext

## 事件返回数据 ​

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| Sheet | Sheet | 表 |
| originValue | String | 原表名 |
| value | String | 现表名 |

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    let eventContext;
    eventContext = await app.Sheets(1).OnRename(data => {
        console.log(data);
        // 取消事件监听
        eventContext.Destroy();
    });
    const sheetName = app.Sheets(1).Name;
    //这里会执行OnRename的回调
    sheetName = 'newName';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    let eventContext;
    eventContext = Application.Sheets(1).OnRename(data => {
        console.log(data);
        // 取消事件监听
        eventContext.Destroy();
    });
    Application.Sheets(1).Name = 'newName';
    //这里会执行OnRename的回调
}
main();
```
 
# 066 API文档 / 数据表 / 监听删除数据表的事件

本页内容

- 说明
- 语法
- 参数
- 返回值
- 事件返回数据
- 事件返回数据示例
- 浏览器环境示例
- 脚本编辑器 示例

# 监听删除数据表的事件 ​

## 说明 ​

为当前数据表添加 Delete 事件,当删除 Sheet 时触发。注意在脚本编辑器中使用时，脚本运行结束就会退出运行，这时可能回调无法被正常触发

## 语法 ​

表达式.OnDelete(Callback)

表达式: Sheet

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Callback | 是 | func | 指定事件发生时的回调方法 ; const eventContext = await Sheet.OnDelete(()=&gt;{ ... }) |

## 返回值 ​

EventContext

## 事件返回数据 ​

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| sheetId | Number | 表的 Id |

## 事件返回数据示例 ​

```javascript
{
    sheetId: 2
}
```

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    await app.Sheets.Add({ Type: 'xlEtDataBaseSheet' });
    let eventContext;
    eventContext = await app.Sheets(1).OnDelete(data => {
        console.log(data);
        // 取消事件监听
        eventContext.Destroy();
    });
    app.Sheets(1).Delete(true);
    //这里会执行OnDelete的回调
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    Application.Sheets.Add({ Type: 'xlEtDataBaseSheet' });
    let eventContext;
    eventContext = Application.Sheets(1).OnDelete(data => {
        console.log(data);
        // 取消事件监听
        eventContext.Destroy();
    });
    Application.Sheets(1).Delete(true);
    //这里会执行OnDelete的回调
}
main();
```
 
# 067 API文档 / 视图 / 新建视图

本页内容

# 新建视图 ​

## 说明 ​

添加视图

## 语法 ​

表达式.Add(Type,Name)

表达式:Views

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Type | 是 | 'Grid'或'Kanban'或'Gallery'或'Form'或’Query‘或‘Gantt’ | 视图类别。 |
| Name | 是 | string | 视图名称 |

## 返回值 ​

[View](/documents/app-integration-dev/guide/dbsheet/Api/View.html), [GridView](/documents/app-integration-dev/guide/dbsheet/Api/GridView.html), [GanttView](/documents/app-integration-dev/guide/dbsheet/Api/GanttView.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const views = await app.Sheets(1).Views;

    const gridView = await views.Add('Grid', '表格视图');
    console.log(await gridView.RowHeight);
    await views.Add('Kanban', '看板视图');
    await views.Add('Gallery', '画册视图');
    await views.Add('Form', '表单视图');
    await views.Add('Query', '查询视图');
    const ganttView = await views.Add('Gantt', '甘特视图');
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const views = Application.Sheets(1).Views;
  const gridView = views.Add('Grid', '表格视图');
  console.log(gridView.RowHeight);
  views.Add('Kanban', '看板视图');
  views.Add('Gallery', '画册视图');
  views.Add('Form', '表单视图');
  views.Add('Query', '查询视图');
  const ganttView = views.Add('Gantt', '甘特视图');
}
main()
```
 
# 068 API文档 / 视图 / 重命名

本页内容

# 重命名 ​

## 说明 ​

重命名、 获取视图名称

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
// 获取视图名称
async function getName() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    const name = await view.Name;
}
// 设置视图名称
async function setName() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    view.Name = '新视图名称';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const view = Application.Sheets(1).Views(1);
  const name = view.Name;
  view.Name = '新视图名称';
}
main()
```
 
# 069 API文档 / 视图 / 复制视图

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 复制视图 ​

## 说明 ​

复制视图

## 语法 ​

表达式.Copy()

表达式:View

## 参数 ​

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    const result = await view.Copy();
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const view = Application.Sheets(1).Views(1);
  const result = view.Copy();
}
main()
```
 
# 070 API文档 / 视图 / 删除视图

本页内容

# 删除视图 ​

## 说明 ​

删除指定视图

## 语法 ​

表达式.Delete()

表达式:View

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    await view.Delete();
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const view = Application.Sheets(1).Views(1);
  view.Delete();
}
main()
```
 
# 071 API文档 / 视图 / 设置个人 / 公共视图

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置个人/公共视图 ​

View.IsPersonal

## 说明 ​

可读写 设置视图为个人视图或公共视图

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function getPublicView() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    const isPersonal = await view.IsPersonal; // 若为个人视图返回true，若不为公共视图返回false
    return !isPersonal;
}

async function setPublicView() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    view.IsPersonal = false; // 设置为公共视图
    view.IsPersonal = true; //  设置为个人视图
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const view = Application.Sheets(1).Views(1);
  const isPublicView = !view.IsPersonal;
  view.IsPersonal = false; // 设置为公共视图
  view.IsPersonal = true; // 设置为个人视图
}
main()
```
 
# 072 API文档 / 视图 / 设置快速访问视图

本页内容

# 设置快速访问视图 ​

View.IsFavView

## 说明 ​

可读写 设置视图是否为快速访问视图

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function getFavView() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    const isFavView = await view.IsFavView; // 若为快速访问视图返回true，若不为快速访问视图返回false
}

async function setFavView() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    view.IsFavView = false; // 取消设置为快速访问视图
    view.IsFavView = true; // 设置为快速访问视图
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function getFavView() {
    const view = Application.Sheets(1).Views(1);
    const isFavView = view.IsFavView; // 若为快速访问视图返回true，若不为快速访问视图返回false
}
function setFavView() {
    const view = Application.Sheets(1).Views(1);
    view.IsFavView = false; // 取消设置为快速访问视图
    view.IsFavView = true; // 设置为快速访问视图
}
main();
```
 
# 073 API文档 / 视图 / 表格视图 / 冻结列数

本页内容

# 冻结列数 ​

GridView.FrozenCols

## 说明 ​

可读写

表格视图的冻结列数

## 返回值 ​

Number

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const gridView = await app.Sheets(1).Views(1);
    console.log(gridView.FrozenCols);
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const gridView = Application.Sheets(1).Views(1);
    console.log(gridView.FrozenCols);
}
main();
```
 
# 074 API文档 / 视图 / 表格视图 / 行高

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 行高 ​

GridView.RowHeight

## 说明 ​

可读写

表格视图的行高，注意：设置值时为以下之一：'Short','Medium','Tall','ExtraTall' .

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const gridView = await app.Sheets(1).Views(1);
    console.log(gridView.RowHeight);
    gridView.RowHeight = 'Tall';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const gridView = Application.Sheets(1).Views(1);
    console.log(gridView.RowHeight);
    gridView.RowHeight = 'Tall';
}
main();
```
 
# 075 API文档 / 视图 / 甘特视图 / 开始日期字段

本页内容

# 开始日期字段 ​

## 说明 ​

可读写

开始日期字段，值为字段 ID.

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const ganttView = await app.Sheets(1).Views(1);
    console.log(ganttView.BeginField);
    ganttView.BeginField = 's';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const ganttView = Application.Sheets(1).Views(1);
    console.log(ganttView.BeginField);
    ganttView.BeginField = 's';
}
main();
```
 
# 076 API文档 / 视图 / 甘特视图 / 结束日期字段

本页内容

# 结束日期字段 ​

## 说明 ​

可读写

结束日期字段，值为字段 ID.

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const ganttView = await app.Sheets(1).Views(1);
    console.log(ganttView.EndField);
    ganttView.EndField = 's';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const ganttView = Application.Sheets(1).Views(1);
    console.log(ganttView.EndField);
    ganttView.EndField = 's';
}
main();
```
 
# 077 API文档 / 视图 / 甘特视图 / 时间线颜色

本页内容

# 时间线颜色 ​

GanttView.TimelineColor

## 说明 ​

可读写

时间线颜色，值为Hex格式的RGB颜色值，如：#FF0000。

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const ganttView = await app.Sheets(1).Views(1);
    console.log(ganttView.TimelineColor);
    ganttView.TimelineColor = '#97E4E4';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const ganttView = Application.Sheets(1).Views(1);
    console.log(ganttView.TimelineColor);
    ganttView.TimelineColor = '#97E4E4';
}
main();
```
 
# 078 API文档 / 视图 / 甘特视图 / 工时统计

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 工时统计-忽略节假日 ​

GanttView.IsOnlyWorkDay

## 说明 ​

可读写

是否忽略节假日。

## 返回值 ​

boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const ganttView = await app.Sheets(1).Views(1);
    console.log(ganttView.IsOnlyWorkDay);
    ganttView.IsOnlyWorkDay = false;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const ganttView = Application.Sheets(1).Views(1);
    console.log(ganttView.IsOnlyWorkDay);
    ganttView.IsOnlyWorkDay = true;
}
main();
```
 
# 079 API文档 / 视图 / 甘特视图 / 定位上一页

本页内容

- 说明
- 返回值
- 浏览器环境示例

# 时间线操作-定位上一页 ​

GanttViewUI.PrevPage

## 说明 ​

跳转到上一页

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function today() {
    await instance.ready();
    const app = instance.Application;
    const GanttViewUI = await app.Window.GanttViewUI
    GanttViewUI.PrevPage()
}
```
 
# 080 API文档 / 视图 / 甘特视图 / 定位至今天

本页内容

# 时间线操作-定位至今天 ​

GanttViewUI.Today(方法)

## 说明 ​

跳转到今天

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function today() {
    await instance.ready();
    const app = instance.Application;
    const GanttViewUI = await app.Window.GanttViewUI
    GanttViewUI.Today()
}
```
 
# 081 API文档 / 视图 / 甘特视图 / 设置折叠

本页内容

# 设置折叠 ​

GanttViewUI.GanttGridFold

## 说明 ​

可读写 设置甘特图字段显示是否折叠

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function today() {
    await instance.ready();
    const app = instance.Application;
    const GanttViewUI = await app.Window.GanttViewUI
    // 设置折叠
    GanttViewUI.GanttGridFold = true
}
```
 
# 082 API文档 / 视图 / 查询视图 / 设置查询条件

本页内容

# 设置查询条件 ​

QueryView.QueryFields

## 说明 ​

可读写

查询视图的查询条件配置数组，可以将数组设置到QueryFields属性，查询条件的数据结构如下

javascript

```javascript
{
conditionCanBlank: false, // 是否必填
customPrompt: "", // 自定义提示语
enableScanCodeToInput: false,  // 是否允许扫码输入
fieldId: "s",  // 字段ID
needSecondCheck: false,  // 电话字段时是否校验号码
op: "Equals" // 匹配方式，参看下面说明
}
```

根据字段类型支持不同的匹配方式 文本/邮箱/URL/地址/级联：Intersected，Contains，Equals 日期： Intersected，GreaterEquAndLessEqu，Equals 时间： Equals 数字/货币/百分比/最后修改时间/等级/进度/创建时间： GreaterEquAndLessEqu， Equals 身份证/电话/自动编号：Intersected，Equals 复选框/单选项/多选项/联系人/创建人/最后修改人/双向关联/单向关联/父记录：Intersected

## 返回值 ​

Array

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    console.log(view.QueryFields);
    // 添加查询条件
    view.QueryFields = [{
        conditionCanBlank: false, // 是否必填
        customPrompt: "", // 自定义提示语
        enableScanCodeToInput: false,  // 是否允许扫码输入
        fieldId: "s",  // 字段ID
        needSecondCheck: false,  // 电话字段时是否校验号码
        op: "Equals" // 匹配方式，参看下面说明
    }]
    // 使用手机验证码
    view.QueryFields = [{
        conditionCanBlank: false,
        customPrompt: "",
        enableScanCodeToInput: false,
        fieldId: "s",
        needSecondCheck: true
        op: "Equals"
    }]
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const view = Application.Sheets(1).Views(1);
    console.log(view.QueryFields);
}
main();
```
 
# 083 API文档 / 视图 / 查询视图 / 设置背景图

本页内容

# 设置背景图 ​

QueryView.BackgroundImage

## 说明 ​

可读写

查询视图的背景图，注意：可以设置为 url/base64

## 返回值 ​

Attachment

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    view.BackgroundImage = "https://kdocs-om.wpscdn.cn/om/image.png"
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const view = Application.Sheets(1).Views(1);
    view.BackgroundImage = "https://kdocs-om.wpscdn.cn/om/image.png"
}
main();
```
 
# 084 API文档 / 视图 / 日历视图 / 时间线颜色

本页内容

# 时间线颜色 ​

CalendarView.TimelineColor(属性)

## 说明 ​

可读写

事件线颜色，值为Hex格式的RGB颜色值，如：#FF0000。

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    console.log(view.TimelineColor);
    view.TimelineColor = '#97E4E4';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const view = Application.Sheets(1).Views(1);
    console.log(view.TimelineColor);
    view.TimelineColor = '#97E4E4';
}
main();
```
 
# 085 API文档 / 视图 / 日历视图 / 开始日期字段

本页内容

# 开始日期字段 ​

CalendarView.BeginField(属性)

## 说明 ​

可读写

开始日期字段，值为字段 ID.

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    console.log(view.BeginField);
    view.BeginField = 's';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const view = Application.Sheets(1).Views(1);
    console.log(view.BeginField);
    view.BeginField = 's';
}
main();
```
 
# 086 API文档 / 视图 / 日历视图 / 结束日期字段

本页内容

# 结束日期字段 ​

CalendarView.EndField(属性)

## 说明 ​

可读写

结束日期字段，值为字段 ID.

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    console.log(view.EndField);
    view.EndField = 's';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const view = Application.Sheets(1).Views(1);
    console.log(view.EndField);
    view.EndField = 's';
}
main();
```
 
# 087 API文档 / 视图 / 日历视图 / 标题设置

本页内容

# 标题设置 ​

CalendarView.TitleField(属性)

## 说明 ​

可读写

日历视图设置标题字段，值为字段 ID.

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const CalendarView = await app.Sheets(1).Views(1);
    console.log(CalendarView.TitleField);
    CalendarView.TitleField = 's';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const CalendarView = Application.Sheets(1).Views(1);
    console.log(CalendarView.TitleField);
    CalendarView.TitleField = 's';
}
main();
```
 
# 088 API文档 / 视图 / 监听增加视图的事件

本页内容

- 说明
- 语法
- 参数
- 返回值
- 事件返回数据
- 浏览器环境示例
- 脚本编辑器 示例

# 监听增加视图的事件 ​

Views.OnCreate(方法)

## 说明 ​

为 Views 添加 Create 事件,当添加 View 时触发。注意在脚本编辑器中使用时，脚本运行结束就会退出运行，这时可能回调无法被正常触发

## 语法 ​

表达式.OnCreate(Callback)

表达式: Views

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Callback | 是 | func | 指定事件发生时的回调方法 ; const eventContext = await Views.OnCreate(()=&gt;{ ... }) |

## 返回值 ​

EventContext

## 事件返回数据 ​

View

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    let eventContext;
    eventContext = await app.Sheets(1).Views.OnCreate(data => {
        console.log(data);
        // 取消事件监听
        eventContext.Destroy();
    });

    await app.Sheets(1).Views.Add('Grid', '表格视图');
    //这里会执行OnCreate的回调
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    let eventContext;
    eventContext = Application.Sheets(1).Views.OnCreate(data => {
        console.log(data);
        // 取消事件监听
        eventContext.Destroy();
    });
    Application.Sheets(1).Views.Add('Grid', '表格视图');
    //这里会执行OnCreate的回调
}
main();
```
 
# 089 API文档 / 视图 / 监听删除视图的事件

本页内容

# 监听删除视图的事件 ​

View.OnDelete(方法)

## 说明 ​

为当前视图添加 Delete 事件,当删除 View 时触发。注意在脚本编辑器中使用时，脚本运行结束就会退出运行，这时可能回调无法被正常触发

## 语法 ​

表达式.OnDelete(Callback)

表达式: View

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Callback | 是 | func | 指定事件发生时的回调方法 ; const eventContext = await View.OnDelete(()=&gt;{ ... }) |

## 返回值 ​

EventContext

## 事件返回数据 ​

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| sheetId | Number | 表的 Id |
| viewId | String | 视图的 Id |

## 事件返回数据示例 ​

```javascript
{
    sheetId: 1
    viewId: 'B'
}
```

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    let eventContext;
    eventContext = await app
        .Sheets(1)
        .Views(1)
        .OnDelete(data => {
            console.log(data);
            // 取消事件监听
            eventContext.Destroy();
        });
    app.Sheets(1).Views(1).Delete();
    //这里会执行OnDelete的回调
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    let eventContext;
    eventContext = Application.Sheets(1)
        .Views(1)
        .OnDelete(data => {
            console.log(data);
            // 取消事件监听
            eventContext.Destroy();
        });
    Application.Sheets(1).Views(1).Delete();
    //这里会执行OnDelete的回调
}
main();
```
 
# 090 API文档 / 视图 / 监听重命名视图的事件

本页内容

# 监听重命名视图的事件 ​

View.OnRename(方法)

## 说明 ​

为当前视图添加 Rename 事件,当被修改 Name 时触发。注意在脚本编辑器中使用时，脚本运行结束就会退出运行，这时可能回调无法被正常触发

## 语法 ​

表达式.OnRename(Callback)

表达式: View

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Callback | 是 | func | 指定事件发生时的回调方法 ; const eventContext = await View.OnRename(()=&gt;{ ... }) |

## 返回值 ​

EventContext

## 事件返回数据 ​

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| View | View | 视图 |
| originValue | String | 原始值 |
| value | String | 修改后的值 |

## 事件返回数据示例 ​

View, originValue, value

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    let eventContext;
    eventContext = await app
        .Sheets(1)
        .Views(1)
        .OnRename(data => {
            console.log(data);
            // 取消事件监听
            eventContext.Destroy();
        });
    const sheetName = app.Sheets(1).Views(1).Name;
    sheetName = 'newName';
    //这里会执行OnRename的回调
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    let eventContext;
    eventContext = Application.Sheets(1)
        .Views(1)
        .OnRename(data => {
            console.log(data);
            // 取消事件监听
            eventContext.Destroy();
        });
    Application.Sheets(1).Views(1).Name = 'newName';
    //这里会执行OnRename的回调
}
main();
```
 
# 091 API文档 / 记录 / 插入记录

本页内容

# 插入记录 ​

## 说明 ​

插入新的记录，在指定行记录之前或之后插入

## 语法 ​

表达式.Add(Before,After,Count)

表达式:RecordRange

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Before | 否 | Number/String | 在行记录之前添加，Number时指定插入时索引，String时指定插入ID |
| After | 否 | Number/String | 在行记录之后添加，Number时指定插入时索引，String时指定插入ID |
| Count | 否 | Number | 一次插入N条数据，未指定时插入1条 |

## 返回值 ​

Self

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  // 在第1行，向上方添加10条记录
  const range = await app.ActiveView.RecordRange.Add(1, null, 10)
  
  const template = ["商品"]
  const range1 = []
  // 给1-10行赋值
  for (let i = 1; i < 11; i++ ) {
    range1.push([template[0]+i,i])
  }
  range.Value = range1

}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const range = Application.ActiveView.RecordRange.Add(31, null, 5)
  // 将插入的5条记录的名称字段 初始化为 “名称”
  range.Item(undefined, "@名称").Value = "名称"
}
main()
```
 
# 092 API文档 / 记录 / 删除记录

本页内容

# 删除记录 ​

## 说明 ​

删除某行记录

## 语法 ​

表达式：Delete(Index)

表达式：Records

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Index | 是 | number/string | 索引从1开始/记录ID |

## 返回值 ​

ApiResult

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    // 删除视图上的第一条数据
    app.Sheets(1).Views(1).Records.Delete(1)
    // 在第10行，向上方添加2条记录
    const records = await app.Sheets(1).Views(1).Records.Add(10, undefined, 2) 
    // 删除刚插入的第一条数据
    records.Delete(1)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    // 删除视图上的第一条数据
    Application.Sheets(1).Views(1).Records.Delete(1)
    // 在第10行，向上方添加2条记录
    const records = Application.Sheets(1).Views(1).Records.Add(10, undefined, 2) 
    // 删除刚插入的第一条数据
    records.Delete(1)
 }
main()
```
 
# 093 API文档 / 记录 / 查看记录

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 查看记录 ​

## 说明 ​

获取指定索引行的记录信息

## 语法 ​

表达式：Item(Index)

表达式：Records

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Index | 是 | number/string | 索引从1开始/记录ID |

## 返回值 ​

Record

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const record = await app.Sheets(1).Views(1).Records.Item(10)
    // 返回字段的文本表示
    console.log(await record.Text)
    // 返回字段的值
    console.log(await record.Value)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const record = Application.Sheets(1).Views(1).Records.Item(10)
    // 返回字段的文本表示
    console.log(record.Text)
    // 返回字段的值
    console.log(record.Value)
 }
main()
```
 
# 094 API文档 / 记录 / 筛选记录

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 筛选记录 ​

## 说明 ​

筛选符合条件的记录

## 语法 ​

表达式.Condition(Filters,FilterOp)

表达式:RecordRange

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Filters | 是 | Filter[] | 筛选数据的条件组, 条件组里每个Filter，可以包含多个筛选条件 |
| FilterOp | 否 | "And"/"Or" | 筛选数据条件组之间的关系，是同时满足还是只需要满足一条，默认值为And |

Filter数据结构：

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Criterias | 是 | [Criteria](/documents/app-integration-dev/guide/dbsheet/Api/Criteria.html)[] | 筛选条件项组 |
| Op | 否 | "And"/"Or" | 筛选条件项组之间的关系，是同时满足还是只需要满足一条，默认值为And |

[Criteria](/documents/app-integration-dev/guide/dbsheet/Api/Criteria.html) 中筛选规则（大小写不敏感）：

| 枚举值 | 描述 |
| --- | --- |
| Equals | 等于 |
| NotEqu | 不等于 |
| Greater | 大于 |
| GreaterEqu | 大等于 |
| Less | 小于 |
| LessEqu | 小等于 |
| GreaterEquAndLessEqu | 介于（取等） |
| LessOrGreater | 不介于（不取等） |
| BeginWith | 开头是 |
| EndWith | 结尾是 |
| Contains | 包含 |
| NotContains | 不包含 |
| Intersected | 指定值 |
| Empty | 为空 |
| NotEmpty | 不为空 |

各筛选规则独立地限制了values数组内最多允许填写的元素数，当values内元素数超过阈值时，该筛选规则将失效。“为空、不为空”不允许填写元素；“介于”允许最多填写2个元素；“指定值”允许填写65535个元素；其他规则允许最多填写1个元素 values[]数组内的元素为字符串时，表示文本匹配。目前还支持对日期进行动态筛选，此时values[]内的元素需以结构体的形式给出：

```javascript
const dateValue = {"dynamicType": "lastMonth","type": "DynamicSimple"}
Criteria("@日期", "Equals", [dateValue])
```

上述示例对应的筛选条件为“等于上一个月”。 要使用日期动态筛选，values[]内的结构体需要指定"type": "DynamicSimple"，当"op"为"equals"时，"dynamicType"可以为如下的值（大小写不敏感）：

| 枚举值 | 描述 |
| --- | --- |
| today | 今天 |
| yesterday | 昨天 |
| tomorrow | 明天 |
| last7Days | 最近7天 |
| last30Days | 最近30天 |
| thisWeek | 本周 |
| lastWeek | 上周 |
| nextWeek | 下周 |
| thisMonth | 本月 |
| lastMonth | 上月 |
| nextMonth | 次月 |

当"op"为"greater"或"less"时，"dynamicType"只能是昨天、今天或明天。

对不同字段类型，values会有不同的用法 联系人字段:

```javascript
// value是一个结构体，指定type为 Contact, value 为用户id
const dateValue = {"type":"Contact", value:"user id"}
```

单/多选项字段:

```javascript
// value是一个结构体，指定type为 SelectItem, value 为选项的ID
const dateValue = {"type":"SelectItem", value:"B"}
```

## 返回值 ​

RecordRange

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  // 创建筛选条件criteria
  const Criteria = app.Criteria 
  const criterias = []
  criterias.push(await Criteria("@名称",  "Intersected", ["1"]))
  // 创建筛选列表filters
  const filters = []
  const filter = {Criterias: criterias, Op: "AND"}
  filters.push(filter)
  // 筛选记录
  const res = await app.ActiveSheet.Views(1).RecordRange.Condition(filters, "AND")
  console.log(res)
  // 操作记录，返回第一个筛选结果的记录ID
  await res.Item(1).Id
  // 删除记录
  await res.Item(1).Delete()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  // 创建筛选条件criteria
    const critera1 = Criteria("@名称", "Intersected", ["1", "999", "aaaa"])
    const critera2 = Criteria("@数量", "Equals", ["1"])
    const criterias = []
    criterias.push(critera1)
    criterias.push(critera2)
    // 创建filters
    const filters = []
    const filter = { Criterias: criterias, Op: "OR" }
    filters.push(filter)
    const res = Application.ActiveSheet.Views(1).RecordRange.Condition(filters, "AND")
    console.log(res.Value)
}
main()
```
 
# 095 API文档 / 记录 / 选中记录

本页内容

# 选中记录 ​

## 说明 ​

选中某个区域

## 语法 ​

表达式.Select()

表达式：Record

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const record = app.Sheets(1).Views(1).Records(10)
    record.Select() // 选中第10行的记录
    
    const info = app.Sheets(1).Views(1).Records(10, 1)
    await info.Select() // 选中第10行，第一个字段
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const record = Application.Sheets(1).Views(1).Records(10)
    record.Select() // 选中第10行的记录
    
    const info = Application.Sheets(1).Views(1).Records(10, 1)
    info.Select() // 选中第10行，第一个字段
}
main()
```
 
# 096 API文档 / 记录 / 展开记录

本页内容

# 展开记录 ​

## 说明 ​

当前窗口下展开记录，显示详情信息

## 语法 ​

表达式.DisplayRecord(RecodId)

表达式: Window

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| RecodId | 否 | String | 展开的记录ID |

## 返回值 ​

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    await app.Window.DisplayRecord("B");
}
```
 
# 097 API文档 / 记录 / 关闭记录

本页内容

- 说明
- 语法
- 返回值
- 浏览器环境示例

# 关闭当前展开的记录 ​

## 说明 ​

关闭当前展开的记录

## 语法 ​

表达式.HiddenAllRecord()

表达式: Window

## 返回值 ​

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    await app.Window.HiddenAllRecord();
}
```
 
# 098 API文档 / 记录 / 设置单元格内容

本页内容

- 说明
- 返回值
    - 这里返回的文本根据区域会有三种情况,RecordRange 选择器参看RecordRange的说明
    - 设置值的时候，也有三种设置值的方式
    - 不同字段类型设置Value时的数据结构
    - 部分不支持设置值的字段类型
- 浏览器环境示例
- 脚本编辑器 示例

# 设置单元格内容 ​

RecordRange.Value

## 说明 ​

可读写

RecordRange的值，读取和设置记录数据

## 返回值 ​

DbCellValue

### 这里返回的文本根据区域会有三种情况,RecordRange 选择器参看[RecordRange](/documents/app-integration-dev/guide/dbsheet/Api/RecordRange.html)的说明 ​

1、指定的区域为一条记录一个字段，返回单个值（值的类型看具体字段）

```javascript
Application.ActiveSheet.RecordRange(1,1).Value
```

2、指定的区域为一条记录的多个字段，返回值为二维数组

```javascript
Application.ActiveSheet.RecordRange(1,[1,2]).Value
```

3、指定的区域为多条记录的多个字段，返回值为二维数组

```javascript
Application.ActiveSheet.RecordRange([1,2],[1,2]).Value
```

### 设置值的时候，也有三种设置值的方式 ​

1、设置为单个值,将这个值设置到整个区域的指定单元格

```javascript
Application.ActiveSheet.RecordRange([1,2],[1,2]).Value = "1"
```

2、传入一维数组，目标是一条记录，则将数组设置到目标区域，如果目标是多条记录，则会将相同的数据设置到所有记录。

```javascript
Application.ActiveView.RecordRange([1,2],[1,2]).Value = ["1","2"]
```

3、传入二维数组，如果二维数组的长度为1，目标是一条记录，则将二维数组[0]设置到目标区域，如果目标是多条记录，则会将相同的数据设置到所有记录。如果二维数组为 M x N，则按顺序传入到目标区域。

```javascript
Application.ActiveView.RecordRange([1,2],[1,2]).Value = [["1","2"],["3","4"]]

// 数组长度为1时等价于 [["1","2"],["1","2"]]
Application.ActiveView.RecordRange([1,2],[1,2]).Value = [["1","2"]]
```

### 不同字段类型设置Value时的数据结构 ​

地址字段类型: 通过DBCellValue() 生成字段的数据

```javascript
Application.Sheets(1).Views(2).RecordRange(1, "@地址").Value = DBCellValue({districts:["广东省","珠海市","香洲区"],detail:"前岛环路xxxx号"})
```

级联字段类型：通过DBCellValue() 生成字段的数据

```javascript
Application.Sheets(1).Views(2).RecordRange(1, "@级联选项").Value = DBCellValue({districts:["广东省","珠海市","香洲区"]})
```

超链接字段类型：

```javascript
Application.RecordRange(1, "@超链接").Value = Application.DBCellValue({address:"wps.cn", display:"wps"})
```

关联字段类型: 参数传入关联的记录id

```javascript
const DBCellValue = Application.DBCellValue
Application.Sheets(1).Views(2).RecordRange(1, "@关联：数据表").Value = DBCellValue(["b","V"])
```

多选项类型：

```javascript
Application.Sheets(1).Views(2).RecordRange([5,6], 4).Value =[[DBCellValue(["未开始","进行中"])], DBCellValue(["进行中"])]
```

联系人字段类型：直接传入联系人的id，如果有多个联系人，可用","进行分割

```javascript
Application.Sheets(1).Views(2).RecordRange([5,6], "@联系人").Value = "238777563"
```

图片与附件字段：可以传入包含 URL/base64编码的图片/云文档 的数组，支持多个附件。 注意：由于脚本有运行时长限制，附件较大/或者较多时会导致超时，设置失败

```javascript
Application.Sheets(1).Views(2).RecordRange(1, "@图片和附件").Value = DBCellValue([{fileData: url/base64, fileName: ""}])
```

如果要在原来的附件上，新增其它附件可以先读出来的数组增加后再重新设置

```javascript
const range = Application.Sheets(1).Views(2).RecordRange(1, "@图片和附件")
const dbCellValue = range.Value
const attments = dbCellValue.Value
attments.push({fileData: url/base64, fileName: ""})
range.Value = oldValue
```

其它字段类型可以直接使用字符串，日期和时间类型必须符合日期的格式的字符串

### 部分不支持设置值的字段类型 ​

```javascript
DbSheetFieldType.Formula, // 公式字段
DbSheetFieldType.Lookup, // 引用字段
DbSheetFieldType.CreatedBy, // 创建者字段
DbSheetFieldType.Note, // 富文本字段
DbSheetFieldType.SearchLookup, // 查找引用字段
DbSheetFieldType.Button,  // 按钮字段 
DbSheetFieldType.LastModifiedBy, // 最近修改者字段
DbSheetFieldType.CreatedTime, // 创建时间字段
DbSheetFieldType.LastModifiedTime, //最后修改时间字段
DbSheetFieldType.AutoNumber, // 自动编号 
DbSheetFieldType.Automations, // 自动任务
```

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
   const app = instance.Application;
   app.Sheets(1).Views(2).RecordRange([5,6], 1).Value = "1111111"
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   Application.Sheets(1).Views(2).RecordRange([5,6], 1).Value = "1111111"
}
main()
```
 
# 099 API文档 / 记录 / 设置单元格的字体颜色

本页内容

# 设置单元格的字体颜色 ​

RecordRange.Font

## 说明 ​

可读写

返回当前RecordRange首个单元格的字体属性[Font](/documents/app-integration-dev/guide/dbsheet/Api/Font.html)

## 返回值 ​

[Font](/documents/app-integration-dev/guide/dbsheet/Api/Font.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const recordRange = await app.Sheet(1).RecordRange(1);
    const font = await recordRange.Font
    font.Color = "#ff00ff"
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const recordRange = Application.Sheet(1).RecordRange(1);
    const font = recordRange.Font
    font.Color = "#ff00ff"
}
main()
```
 
# 100 API文档 / 记录 / 设置单元格的填充颜色

本页内容

# 设置单元格的填充颜色 ​

RecordRange.Interior

## 说明 ​

可读 返回当前RecordRange首个单元格的填充属性[Interior](/documents/app-integration-dev/guide/dbsheet/Api/Interior.html)

## 返回值 ​

[Interior](/documents/app-integration-dev/guide/dbsheet/Api/Interior.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const recordRange = await app.Sheet(1).RecordRange(1);
    const Interior = await recordRange.Interior
    Interior.Color = "#ff00ff"
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const recordRange = Application.Sheet(1).RecordRange(1);
    const Interior = recordRange.Interior
    Interior.Color = "#ff00ff"
}
main()
```
 
# 101 API文档 / 记录 / 监听修改记录的事件

本页内容

- 说明
- 语法
- 参数
- 返回值
- 事件返回数据
- 浏览器环境示例
- 脚本编辑器 示例

# 监听修改记录的事件 ​

RecordRange.OnUpdate(方法)

## 说明 ​

为 RecordRange 添加 Update 事件,当更新 RecordRange 时触发。注意在脚本编辑器中使用时，脚本运行结束就会退出运行，这时可能回调无法被正常触发

## 语法 ​

表达式: OnUpdate(Callback)

表达式: RecordRange

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Callback | 是 | func | 指定事件发生时的回调方法 ; const eventContext = await RecordRange.OnUpdate(()=&gt;{ ... }) |

## 返回值 ​

EventContext

## 事件返回数据 ​

[RecordRange]

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    let eventContext;
    eventContext = await app
        .Sheets(1)
        .Views(1)
        .RecordRange(1)
        .OnUpdate(data => {
            console.log(data);
            // 取消事件监听
            eventContext.Destroy();
        });
    app.Sheets(1).Views(1).RecordRange(1).Value = ['名称111', 4, '选项1'];
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    let eventContext;
    eventContext = Application.Sheets(1)
        .Views(1)
        .RecordRange(1)
        .OnUpdate(data => {
            console.log(data);
            // 取消事件监听
            eventContext.Destroy();
        });
    Application.Sheets(1).Views(1).RecordRange(1).Value = ['名称111', 4, '选项1'];
}
main();
```
 
# 102 API文档 / 记录 / 监听删除记录的事件

本页内容

- 说明
- 语法
- 参数
- 返回值
- 事件返回数据
- 事件返回数据示例
- 浏览器环境示例
- 脚本编辑器 示例

# 监听删除记录的事件 ​

RecordRange.OnDeleteRecord(方法)

## 说明 ​

为 RecordRange 添加 DeleteRecord 事件,当删除 RecordRange 时触发。注意在脚本编辑器中使用时，脚本运行结束就会退出运行，这时可能回调无法被正常触发

## 语法 ​

表达式: OnDeleteRecord(Callback)

表达式: RecordRange

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Callback | 是 | func | 指定事件发生时的回调方法 ; const eventContext = await RecordRange.OnDeleteRecord(()=&gt;{ ... }) |

## 返回值 ​

EventContext

## 事件返回数据 ​

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| sheetId | Number | 表的 Id |
| recordIds | Array | 记录集合的 Ids |

## 事件返回数据示例 ​

```javascript
{
    recordIds: ['A','C']
    sheetId: 1
}
```

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    let eventContext;
    eventContext = await app
        .Sheets(1)
        .Views(1)
        .RecordRange(1)
        .OnDeleteRecord(data => {
            console.log(data);
            // 取消事件监听
            eventContext.Destroy();
        });
    app.Sheets(1).Views(1).RecordRange(1).Delete();
    //这里会执行OnDeleteRecord的回调
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    let eventContext;
    eventContext = Application.Sheets(1)
        .Views(1)
        .RecordRange(1)
        .OnDeleteRecord(data => {
            console.log(data);
            // 取消事件监听
            eventContext.Destroy();
        });
    Application.Sheets(1).Views(1).RecordRange(1).Delete();
    //这里会执行OnDeleteRecord的回调
}
main();
```
 
# 103 API文档 / 字段 / 新增字段

本页内容

# 新增字段 ​

## 说明 ​

往表中新增新的字段

## 语法 ​

表达式.AddField(FieldDescriptor, Index)

表达式:FieldDescriptors

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| FieldDescriptor | 是 | FieldDescriptor | 字段属性 |
| Index | 否 | string/number | Index为string时表示字段ID，number时表示字段索引，插入位置，未指定时插入到末尾 |

## 返回值 ​

ApiResult

## jsApi 示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const desc = await app.FieldDescriptor("Rating","等级字段")
  desc.MaxRating = 2
  await app.Sheets(1).FieldDescriptors.AddField(desc, 1)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const desc = FieldDescriptor("Rating","等级字段")
  desc.MaxRating = 2
  await Application.Sheets(1).FieldDescriptors.AddField(desc, 1)
}
main()
```
 
# 104 API文档 / 字段 / 复制字段

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器示例

# 复制字段 ​

## 说明 ​

复制当前字段到指定位置

## 语法 ​

表达式.Copy(Before,After)

表达式:Field

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Before | 否 | String/Number | 在Before之前插入复制字段 |
| After | 否 | String/Number | 在After之后插入复制字段 |

## 返回值 ​

Self

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const field = await app.Sheets(1).Views(1).Fields(1)
 field.Copy(3)
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 const field = Application.Sheets(1).Views(1).Fields(1)
 field.Copy(3)
}
main()
```
 
# 105 API文档 / 字段 / 隐藏 / 显示字段

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器示例

# 隐藏/显示字段 ​

Field.Visible(属性)

## 说明 ​

可读写

视图的字段是否可见

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const field = await app.Sheets(1).Views(1).Fields(1)
 field.Visible = false
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 const field = Application.Sheets(1).Views(1).Fields(1)
 const visible = field.Visible
 console.log(visible)
}
main()
```
 
# 106 API文档 / 字段 / 删除字段

本页内容

- 说明
- 语法
- 参数
- 返回值
- jsApi 示例
- 脚本编辑器 示例

# 删除字段 ​

## 说明 ​

删除字段

## 语法 ​

表达式.Delete(RemoveReversedLink)

表达式:FieldDescriptor

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| RemoveReversedLink | 否 | Boolean |  |

## 返回值 ​

Boolean

## jsApi 示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
   const app = instance.Application;
   const field = await WPSOpenApi.Application.Sheets(1).FieldDescriptors(2)
   field.Delete()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const app = Application;
   const field = Application.Sheets(1).FieldDescriptors(2)
   field.Delete()
}
main()
```
 
# 107 API文档 / 字段 / 移动字段

本页内容

# 移动字段 ​

## 说明 ​

移动字段到指定位置

## 语法 ​

表达式.Move(Before,After)

表达式:Field

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Before | 否 | String/Number | 移动到Before字段之前 |
| After | 否 | String/Number | 移动到After字段之后 |

## 返回值 ​

ApiResult

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const field = await app.Sheets(1).Views(1).Fields(1)
 field.Move(3)
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 const field = Application.Sheets(1).Views(1).Fields(1)
 field.Move(3)
}
main()
```
 
# 108 API文档 / 字段 / 设置字段名

本页内容

# 设置字段名称 ​

FieldDescriptor.Name(属性)

## 说明 ​

可读写

字段名称

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const field = await app.Sheets(1).FieldDescriptors(2)
  field.Name = "字段名"
  field.Apply()

  console.log(await field.Name)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const field = Application.Sheets(1).FieldDescriptors(2)
  field.Name = "字段名"
  field.Apply()

  console.log(await field.Name)
}
main()
```
 
# 109 API文档 / 字段 / 设置字段类型

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置字段类型 ​

FieldDescriptor.Type

## 说明 ​

可读写

返回当前字段类型

## 返回值 ​

| 字段类型 | 描述 |
| --- | --- |
| ID | 身份证 |
| Phone | 电话 |
| Email | 电子邮箱 |
| Url | 超链接 |
| Checkbox | 复选框 |
| SingleSelect | 单选项 |
| MultipleSelect | 多选项 |
| Rating | 等级 |
| Complete | 进度条 |
| CellPicture | 单元格图片 |
| Contact | 联系人 |
| Attachment | 附件 |
| Note | 富文本字段，备注 |
| Link | 关联 |
| OneWayLink | 单向关联 |
| Lookup | 引用 |
| Address | 地址,特殊级联字段 |
| Cascade | 级联 |
| Automations | 触发器 |
| AutoNumber | 编号 |
| CreatedBy | 创建者 |
| CreatedTime | 创建时间 |
| LastModifiedBy | 最后修改者 |
| LastModifiedTime | 最后修改时间 |
| Formula | 公式 |

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
   const app = instance.Application;
   const field = await app.Sheets(1).FieldDescriptors(2)
   console.log(await field.Type)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const field = Application.Sheets(1).FieldDescriptors(2)
   console.log(field.Type)
}
main()
```
 
# 110 API文档 / 字段 / 设置默认值

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置默认值 ​

FieldDescriptor.DefaultVal(属性)

## 说明 ​

可读写

设置默认值

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
   const app = instance.Application;
   const field = await WPSOpenApi.Application.Sheets(1).FieldDescriptors(2)
   field.DefaultVal = "1"
   field.Apply()
   console.log(await field.DefaultVal)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const field = Application.Sheets(1).FieldDescriptors(2)
   field.DefaultVal = "1"
   field.Apply()
   console.log(field.DefaultVal)
}
main()
```
 
# 111 API文档 / 字段 / 禁止录入重复值

本页内容

# 禁止录入重复值 ​

FieldDescriptor.ValueUnique(属性)

## 说明 ​

可读写

是否唯一值，禁止重复录入

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const field = await app.Sheets(1).FieldDescriptors(2)
  console.log(await field.ValueUnique)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const field = Application.Sheets(1).FieldDescriptors(2)
  console.log(field.ValueUnique)
}
main()
```
 
# 112 API文档 / 字段 / 日期字段 / 获取字段值

本页内容

# 获取日期类型值 ​

## 说明 ​

获取 日期字段 类型值

## 返回 ​

`string` 类型，如 '2024/12/30 星期一 00:00'

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    // 输出日期值 '2024/12/30 星期一 00:00'
    const value = await app.Sheets(1).Views(1).RecordRange(1, "@日期").Value
    console.log(value)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const value = Application.Sheets(1).Views(1).RecordRange(1, "@日期").Value
    console.log(value)
}
main()
```
 
# 113 API文档 / 字段 / 日期字段 / 设置字段值

本页内容

- 说明
- 浏览器环境示例
- 脚本编辑器 示例

# 设置日期类型值 ​

## 说明 ​

设置 日期字段 类型值，如，"2024/12/25"

设置时，暂不支持带有星期的信息，如 "2024/12/25 星期三" 会报错

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const data = getNowDate()
    // 设置日期 "2024/12/30"
    app.Sheets(1).Views(1).RecordRange(2, "@日期").Value = data
}
// 获取当前日期，格式为 "yyyy:MM:dd"
function getNowDate() {
  const date = new Date()
  return date.getFullYear() + '/' + (date.getMonth() + 1) + '/' + date.getDate()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    // 设置云文档类型
    Application.Sheets(1).Views(1).RecordRange(2, "@日期").Value = '2024/12/30'
}
main()
```
 
# 114 API文档 / 字段 / 日期字段 / 显示星期

本页内容

# 是否显示星期 ​

DateField.IsShowWeek(属性)

## 说明 ​

可读写

字段类型为 日期 时，通过此属性设置是否显示星期

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const Application = instance.Application
  const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@日期")
  const prop = await fieldDescriptor.Date
  prop.IsShowWeek = true
  fieldDescriptor.Apply()
  console.log(await prop.IsShowWeek)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@日期")
  const prop = fieldDescriptor.Date
  prop.IsShowWeek = true
  fieldDescriptor.Apply()
}
main()
```
 
# 115 API文档 / 字段 / 日期字段 / 显示时间

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 显示时间 ​

DateField.IsShowTime(属性)

## 说明 ​

可读写

字段类型为 日期 时，通过此属性可以设置是否显示时间

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const Application = instance.Application
  const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@日期")
  const prop = await fieldDescriptor.Date
  prop.IsShowTime = true
  fieldDescriptor.Apply()
  console.log(await prop.IsShowTime)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@日期")
  const prop = fieldDescriptor.Date
  prop.IsShowTime = true
  fieldDescriptor.Apply()
}
main()
```
 
# 116 API文档 / 字段 / 日期字段 / 面板标记休息日

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 面板标记休息日 ​

DateField.IsShowHoliday(属性)

## 说明 ​

可读写

只对日期字段有效，是否在选择日期的面板上标识节假日

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const Application = instance.Application
  const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@日期")
  const prop = await fieldDescriptor.Date
  prop.IsShowHoliday = true
  fieldDescriptor.Apply()
  console.log(await prop.IsShowHoliday)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@日期")
  const prop = fieldDescriptor.Date
  prop.IsShowHoliday = true
  fieldDescriptor.Apply()
}
main()
```
 
# 117 API文档 / 字段 / 数字字段 / 获取字段值

本页内容

# 获取数字字段类型值 ​

## 说明 ​

获取 数字字段 类型值

## 返回 ​

`number` 类型，如 `100`

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    // 输出数字值 100
    const value = await app.Sheets(1).Views(1).RecordRange(1, "@数字").Value
    console.log(value)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const value = Application.Sheets(1).Views(1).RecordRange(1, "@数字").Value
    console.log(value)
}
main()
```
 
# 118 API文档 / 字段 / 数字字段 / 设置字段值

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置数字字段类型值 ​

## 说明 ​

设置 数字字段 类型值

## 返回值 ​

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    // 设置数字
    app.Sheets(1).Views(1).RecordRange(2, "@数字").Value = 1
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    // 设置数字
    Application.Sheets(1).Views(1).RecordRange(2, "@数字").Value = 1
}
main()
```
 
# 119 API文档 / 字段 / 数字字段 / 显示千位符

本页内容

# 是否显示千位符 ​

NumberField.IsShowThousand(属性)

## 说明 ​

可读写

字段类型为Number时，通过此属性可以快捷的设置是否显示千位符

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const field = await app.Sheets(1).FieldDescriptors("@数字")
  const prop = await field.Number
  prop.IsShowThousand = true
  field.Apply()
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const field = Application.Sheets(1).FieldDescriptors("@数字")
  const prop = field.Number
  prop.IsShowThousand = true
  field.Apply()
 }
main()
```
 
# 120 API文档 / 字段 / 超链接字段 / 获取字段值

本页内容

- 说明
- 返回
    - object结构
- 浏览器环境示例
- 脚本编辑器 示例

# 获取超链接字段类型值 ​

## 说明 ​

获取 超链接字段 类型值

## 返回 ​

`object` 结构，结构如下：

### object结构 ​

| key | 值类型 | 说明 |
| --- | --- | --- |
| address | string | 超链接地址 |
| display | string | 展示的文案 |

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const value = await app.Sheets(1).Views(1).RecordRange(1, "@超链接").Value
    console.log(value)
    /**
     * 输出值：
     * {
     *  address: "https://365.kdocs.cn/", 
     *  display: "跳转文档"
     * }
     */
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const value = Application.Sheets(1).Views(1).RecordRange(1, "@超链接").Value
    console.log(value)
}
main()
```
 
# 121 API文档 / 字段 / 超链接字段 / 设置字段值

本页内容

# 设置超链接类型值 ​

## 说明 ​

设置 超链接字段 类型值

## 语法 ​

```javascript
DBCellValue([{
  address: "",
  display: "",
}])
```

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| address | 是 | string | 超链接地址 |
| display | 是 | string | 展示的文案 |

## 返回值 ​

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    // 设置超链接
    app.Sheets(1).Views(1).RecordRange(2, "@超链接").Value = await Application.DBCellValue({
      address:"wps.cn", 
      display:"wps"
    })
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    // 设置超链接
    Application.Sheets(1).Views(1).RecordRange(2, "@超链接").Value = Application.DBCellValue({
      address:"wps.cn", 
      display:"wps"
    })
}
main()
```
 
# 122 API文档 / 字段 / 超链接字段 / 设置显示样式

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置显示样式 ​

UrlField.HyperLinkText(属性)

## 说明 ​

可读写

只对超链接字段有效，超链接字段设置显示的文本，如果设置了 HyperLinkText ，超链接字段显示样式就会以按钮形式显示

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const field = await app.Sheets(1).FieldDescriptors("@超链接")
  const prop = await field.Url
  // 以按钮形式显示 "Go"
  prop.HyperLinkText = "Go"
  field.Apply()

  // 以超链接形式显示
  prop.HyperLinkText = ""
  field.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const field = Application.Sheets(1).FieldDescriptors("@超链接")
  const prop = field.Url
  // 以按钮形式显示 "Go"
  prop.HyperLinkText = "Go"
  field.Apply()

  // 以超链接形式显示
  prop.HyperLinkText = ""
  field.Apply()
}
main()
```
 
# 123 API文档 / 字段 / 单选项 / 多选项字段 / 获取字段值

本页内容

# 获取单选项/多选项类型值 ​

## 说明 ​

获取 单选项字段 或者 多选项字段 类型值

## 返回 ​

1. 单选项字段类型时，返回 `string` 类型，如 '选项1'
2. 多选项字段类型时，返回`string[]`类型，如 ['选项1', '选项2']

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const value = await app.Sheets(1).Views(1).RecordRange(1, "@单选项").Value
    console.log(value)
    // 输出值：'选项1'
    const value2 = await app.Sheets(1).Views(1).RecordRange(1, "@多选项").Value
    console.log(value)
    // 输出值：['选项1', '选项2']
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const value = Application.Sheets(1).Views(1).RecordRange(1, "@单选项").Value
    console.log(value)
}
main()
```
 
# 124 API文档 / 字段 / 单选项 / 多选项字段 / 设置字段值

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置单选项/多选项类型值 ​

## 说明 ​

设置 单选项字段 或者 多选项字段 类型值

## 返回值 ​

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    // 设置单选项
    app.Sheets(1).Views(1).RecordRange(2, "@单选项").Value = ["选项1"]
    // 设置多选项
    app.Sheets(1).Views(1).RecordRange(2, "@多选项").Value =  [[await Application.DBCellValue(["选项1","选项2","选项3"])]]
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    // 设置单选项
    Application.Sheets(1).Views(1).RecordRange(2, "@单选项").Value = ["选项1"]
    // 设置多选项
    Application.Sheets(1).Views(1).RecordRange(2, "@多选项").Value =  [[Application.DBCellValue(["选项1","选项2","选项3"])]]
}
main()
```
 
# 125 API文档 / 字段 / 单选项 / 多选项字段 / 设置选项

本页内容

# 设置选项 ​

SelectField.Items(属性)

## 说明 ​

可读写

设置多项式字段的可选项

## 返回值 ​

Array

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const filed = await app.Sheets(1).FieldDescriptors("@状态")
  const prop = await filed.Select
  const Items = await prop.Items
  // 加一项
  Items.push({"value":"未开始233", "colorHex":"#0081C2"})
  prop.Items = Items
  filed.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main(){
  const filed = Application.Sheets(1).FieldDescriptors("@状态")
  const prop = field.Select
  const Items = prop.Items
  // 加一项
  Items.push({"value":"未开始233", "colorHex":"#0081C2"})
  prop.Items = Items
  filed.Apply()
}
main()
```
 
# 126 API文档 / 字段 / 单选项 / 多选项字段 / 允许填写时添加选项

本页内容

# 允许填写时添加选项 ​

SelectField.IsAddItemWhenInputting(属性)

## 说明 ​

可读写

当字段类型为单选项或多选项时，可以通过设置 IsAddItemWhenInputting 属性设置允许填写时添加选项。当字段类型不是单选项或多选项时，属性无效。

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
   const app = instance.Application;
   const field = await app.Sheets(1).FieldDescriptors(2)
   const prop = await field.Select
   prop.IsAddItemWhenInputting = true
   field.Apply()
   console.log(await prop.IsAddItemWhenInputting)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const field = Application.Sheets(1).FieldDescriptors(2)
   const prop = field.Select
   prop.AllowAddItemWhenInputting = true
   field.Apply()
   console.log(prop.AllowAddItemWhenInputing)
}
main()
```
 
# 127 API文档 / 字段 / 等级字段 / 获取字段值

本页内容

- 说明
- 返回
- 浏览器环境示例
- 脚本编辑器 示例

# 获取等级类型值 ​

## 说明 ​

获取 等级字段 类型值

## 返回 ​

`number` 类型，如 2

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const value = await app.Sheets(1).Views(1).RecordRange(1, "@等级").Value
    console.log(value)
    // 输出2
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const value = Application.Sheets(1).Views(1).RecordRange(1, "@等级").Value
    console.log(value)
}
main()
```
 
# 128 API文档 / 字段 / 等级字段 / 设置字段值

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置等级类型值 ​

## 说明 ​

设置 等级字段 类型值

## 返回值 ​

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    // 设置等级
    app.Sheets(1).Views(1).RecordRange(2, "@等级").Value = 1
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    // 设置等级
    Application.Sheets(1).Views(1).RecordRange(2, "@等级").Value = 1
}
main()
```
 
# 129 API文档 / 字段 / 等级字段 / 设置等级

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置等级 ​

RatingField.MaxRating(属性)

## 说明 ​

可读写

只对等级字段有效，设置等级字段最大值，取值范围为1-5

## 返回值 ​

Number

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const field = await app.Sheets(1).FieldDescriptors(2)
  const prop = await field.Rating
  prop.MaxRating = 5
  field.Apply()

  console.log(await prop.MaxRating)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const field = Application.Sheets(1).FieldDescriptors(2)
  const prop = await field.Rating
  prop.MaxRating = 5
  field.Apply()

  console.log(prop.MaxRating)
}
main()
```
 
# 130 API文档 / 字段 / 图片和附件字段 / 获取字段值

本页内容

- 说明
- 返回
    - Attachment结构
- 浏览器环境示例
- 脚本编辑器 示例

# 获取图片与附件字段类型值 ​

## 说明 ​

获取 图片与附件字段 类型值

## 返回 ​

[Attachment](/documents/app-integration-dev/guide/dbsheet/Api/Attachment.html) 数据结构，结构如下：

### Attachment结构 ​

| key | 值类型 | 说明 |
| --- | --- | --- |
| FileId | string | 附件字段的Id |
| fileName | string | 附件名称 |
| FileSize | number | 附件大小 |
| FileType | string | 附件类型 |
| ImgSize | string | 图片的尺寸 |
| LinkUrl | string | 附件的链接地址，如果是云文档则返回云文档的地址，如果是图片则返回图片的下载地址 |

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const value = await app.Sheets(1).Views(1).RecordRange(1, "@图片与附件").Value
    console.log(value)
    /**
     * 输出值：
     * {
     *  FileId: "GCN3WSY5ADQAI", 
     *  FileName: "11.png"
     *  FileSize: 10052
     *  FileType: "image/png"
     *  ImgSize: "120*120"
     * }
     */
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const value = Application.Sheets(1).Views(1).RecordRange(1, "@图片与附件").Value
    console.log(value)
}
main()
```
 
# 131 API文档 / 字段 / 图片和附件字段 / 设置字段值

本页内容

# 设置图片与附件类型值 ​

## 说明 ​

设置 图片与附件字段 类型值

## 语法 ​

```javascript
DBCellValue([{
  fileData: "",
  fileName: "",
}])
```

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| fileData | 是 | string | 支持2种类型：`base64` 或 `云文档链接` |
| fileName | 是 | string | 附件名称 |

## 返回值 ​

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    // 设置云文档类型
    app.Sheets(1).Views(1).RecordRange(2, "@图片和附件").Value = await Application.DBCellValue([{
      fileData: "https://kdocs.cn/l/csGRGIzv9PvF",
      fileName: "11.png",
    }])
    // 设置base64
    app.Sheets(1).Views(1).RecordRange(2, "@图片和附件").Value = await Application.DBCellValue([{
      fileData: url/base64,
      fileName: "11.png",
    }])
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    // 设置云文档类型
    Application.Sheets(1).Views(1).RecordRange(2, "@图片和附件").Value = Application.DBCellValue([{
      fileData: "https://kdocs.cn/l/csGRGIzv9PvF",
      fileName: "11.png",
    }])
    // 设置base64
    Application.Sheets(1).Views(1).RecordRange(2, "@图片和附件").Value = Application.DBCellValue([{
      fileData: url/base64,
      fileName: "11.png",
    }])
}
main()
```
 
# 132 API文档 / 字段 / 图片和附件字段 / 设置显示样式

本页内容

# 设置显示样式 ​

AttachmentField.DisplayStyle(属性)

## 说明 ​

可读写

图片和附件字段的显示样式，是以缩略图样式显示还是以列表的样式显示

## 返回值 ​

Enum.DbAttachmentDisplayStyle

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const field = await app.Sheets(1).FieldDescriptors("@图片和附件")
  const prop = await field.Attachment
  prop.DisplayStyle = "Pic"
  field.Apply()
  console.log(await prop.DisplayStyle)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const field = Application.Sheets(1).FieldDescriptors("@图片和附件")
  const prop = field.Attachment
  prop.DisplayStyle = "List"
  field.Apply()
}
main()
```
 
# 133 API文档 / 字段 / 图片和附件字段 / 仅可通过移动端拍摄上传

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 仅可通过移动端拍摄上传 ​

AttachmentField.IsOnlyCameraUpload(属性)

## 说明 ​

可读写

图片和附件字段是否仅可通过移动端拍摄上传

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const field = await app.Sheets(1).FieldDescriptors("@图片和附件")
  const prop = await field.Attachment
  prop.IsOnlyCameraUpload = true
  field.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const field = Application.Sheets(1).FieldDescriptors("@图片和附件")
  const prop = field.Attachment
  prop.IsOnlyCameraUpload = true
  field.Apply()
}
main()
```
 
# 134 API文档 / 字段 / 联系人字段 / 获取字段值

本页内容

# 获取联系人类型值 ​

## 说明 ​

获取 联系人字段 类型值

## 返回 ​

`string`类型，返回多个联系人id时用","进行分割

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const value = await app.Sheets(1).Views(1).RecordRange(1, "@联系人").Value
    console.log(value)
    // 输出 "1557445524,238777563"
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const value = Application.Sheets(1).Views(1).RecordRange(1, "@联系人").Value
    console.log(value)
}
main()
```
 
# 135 API文档 / 字段 / 联系人字段 / 设置字段值

本页内容

# 设置联系人类型值 ​

## 说明 ​

设置 联系人字段 类型值

## 返回值 ​

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    // 设置联系人字段，值为用户id
    app.Sheets(1).Views(1).RecordRange(2, "@联系人").Value = "1557445524"
    // 设置多个联系人字段，用户id用","进行分割
    app.Sheets(1).Views(1).RecordRange(2, "@联系人").Value = "1557445524,238777563"
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    // 设置联系人字段
    Application.Sheets(1).Views(1).RecordRange(2, "@联系人").Value = "1557445524,238777563"
}
main()
```
 
# 136 API文档 / 字段 / 联系人字段 / 插入多个联系人

本页内容

# 插入多个联系人 ​

ContactField.IsSupportMulti(属性)

## 说明 ​

可读写

字段类型为联系人时，通过此属性可以设置是否允许向单元格插入多个联系人

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const Application = instance.Application
    const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@联系人")
    const prop = await fieldDescriptor.Contact
    prop.IsSupportMulti = true
    fieldDescriptor.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@联系人")
    const prop = fieldDescriptor.Contact
    prop.IsSupportMulti = true
    fieldDescriptor.Apply()
}
main()
```
 
# 137 API文档 / 字段 / 联系人字段 / 发送通知

本页内容

# 发送通知 ​

ContactField.IsSupportNotice(属性)

## 说明 ​

可读写

字段类型为联系人时，通过此属性可以设置是否允许向新插入的联系人发送通知

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const Application = instance.Application
    const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@联系人")
    const prop = await fieldDescriptor.Contact
    prop.IsSupportNotice = true
    fieldDescriptor.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@联系人")
    const prop = fieldDescriptor.Contact
    prop.IsSupportNotice = true
    fieldDescriptor.Apply()
}
main()
```
 
# 138 API文档 / 字段 / 关联字段 / 获取字段值

本页内容

# 获取关联字段类型值 ​

## 说明 ​

获取 关联字段 类型值

## 返回 ​

`object` 结构，结构如下：

### object结构 ​

| key | 值类型 | 说明 |
| --- | --- | --- |
| id | string | 关联的引用字段id |
| str | string | 展示的文案 |

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const value = await app.Sheets(1).Views(1).RecordRange(1, "@关联").Value
    console.log(value)
    /**
     * 输出值：
     * {
     *  id: "R", 
     *  str: "数据表2-文本1"
     * }
     */
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const value = Application.Sheets(1).Views(1).RecordRange(1, "@关联").Value
    console.log(value)
}
main()
```
 
# 139 API文档 / 字段 / 关联字段 / 设置字段值

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置关联字段类型值 ​

## 说明 ​

设置 关联字段 类型值

## 返回值 ​

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    // 设置关联字段值，参数为传入关联的记录id
    app.Sheets(1).Views(1).RecordRange(3, "@关联").Value = await Application.DBCellValue(["G", "H"])
    
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    // 设置关联字段值，参数为传入关联的记录id
    Application.Sheets(1).Views(1).RecordRange(2, "@联系人").Value = Application.DBCellValue(["G", "H"])
}
main()
```
 
# 140 API文档 / 字段 / 关联字段 / 关联多条记录

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 关联多条记录 ​

LinkField.IsSupportMultiLinks(属性)

## 说明 ​

可读写

设置或读取 单向关联字段或双向关联字段 是否可以设置允许关联多条记录

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const field = await app.Sheets(1).FieldDescriptors(2)
    const link = await field.Link
    field.IsSupportMultiLinks = true
	field.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const field = Application.Sheets(1).FieldDescriptors(2)
    const link = field.Link
    link.IsSupportMultiLinks = true
	field.Apply() 
}
main()
```
 
# 141 API文档 / 字段 / 关联字段 / 自动关联

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 自动关联 ​

LinkField.IsAutoLink(属性)

## 说明 ​

可读写

设置或读取 单向关联字段 或 双向关联字段 是否自动关联

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const field = await app.Sheets(1).FieldDescriptors(2)
    const prop = await field.Link
    console.log(await prop.IsAutoLink)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const field = Application.Sheets(1).FieldDescriptors(2)
    const prop = field.Link
    console.log(prop.IsAutoLink)
}
main()
```
 
# 142 API文档 / 字段 / 关联字段 / 关联表格ID

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 关联表格ID ​

FieldDescriptor.LinkSheet(属性)

## 说明 ​

可读写

设置或读取单向关联或双向关联字段的关联表格的Id

## 返回值 ​

Number

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    // 插入单向关联字段
    const Application = instance.Application
    const sheet = await Application.Sheets(1)
    const linkField = await Application.FieldDescriptor("OneWayLink","单向关联")
    const prop = await linkField.Link
    prop.LinkSheet = await sheet.StId
    prop.IsAutoLink = false
    sheet.FieldDescriptors.AddField(linkField)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const sheet = Application.Sheets(1)
    const linkField = Application.FieldDescriptor("OneWayLink","单向关联")
    const prop = linkField.Link
    prop.LinkSheet = sheet.StId
    prop.IsAutoLink = false
    sheet.FieldDescriptors.AddField(linkField)
}
main()
```
 
# 143 API文档 / 字段 / 关联字段 / 关联视图ID

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置关联视图ID ​

FieldDescriptor.LinkView(属性)

## 说明 ​

可读写

设置或读取单向关联或双向关联字段的关联视图的id

![alt text](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/FieldDescriptor_LinkView.BeJbQ7Ke.png)

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const field = await app.Sheets(1).FieldDescriptors(2)
    const prop = await field.Link
    prop.LinkView = await app.ActiveView.Id
    field.Apply()
    console.log(await prop.LinkView)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const field = Application.Sheets(1).FieldDescriptors(2)
    const prop = field.Link
    prop.LinkView = ActiveView.Id
    field.Apply()
    console.log(prop.LinkView)
}
main()
```
 
# 144 API文档 / 字段 / 关联字段 / 关系组集合

本页内容

# 设置关系组集合 ​

FieldDescriptor.AutoLinkGroups(属性)

## 说明 ​

可读写

设置或读取 关联字段的关系组集合，如果有多个关系组，则这些关系组是或的关系

![alt text](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/FieldDescriptor_AutoLinkGroups.Dg5N6jJT.png)

## 返回值 ​

AutoLinkGroups

## 浏览器环境示例 ​

javascript

```javascript
// 设置数据表的关联字段的匹配条件
// 匹配到数据表2：数据表2的状态字段 = 数据表的文本字段，如下设置：
async function example() {
  await instance.ready();
  const Application = instance.Application;
  // 数据表
  const sheet = Application.Sheets(1)
  // 数据表的字段
  const fieldId = await sheet.FieldId("文本")
  // 数据表的关联字段
  const linkField = await sheet.FieldDescriptors("@关联")
  const prop = linkField.Link
  const linkGroups = await Application.AutoLinkGroups()
  const group = linkGroups.Add()
  const conditions = group.Conditions
  // 关联的数据表2
  const linkSheet = await Application.Sheets(2)
  // 关联数据表2的字段
  const linkSheet_fieldId = await linkSheet.FieldId("状态")
  // 生成匹配条件
  conditions.Add(linkSheet_fieldId, [fieldId], "Field", "Equals")
  // 设置关联字段的匹配条件
  prop.AutoLinkGroups = linkGroups
  // 设置自动关联
  prop.IsAutoLink = true
  linkField.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main(){
  const sheet = Application.Sheets(1)
  const fieldId = sheet.FieldId("状态")
  // 获取关联字段
  const linkField = sheet.FieldDescriptors("@关联")
  const prop = linkField.Link
  const linkGroups = Application.AutoLinkGroups()
  const conditions = linkGroups.Conditions
  // 关联的数据表2
  const linkSheet = Application.Sheets(2)
  // 关联数据表2的字段
  const linkSheet_fieldId = linkSheet.FieldId("状态")
  // 生成匹配条件
  conditions.Add(linkSheet_fieldId, [fieldId], "Field", "Equals")
  // 设置关联字段的匹配条件
  prop.AutoLinkGroups = linkGroups
  // 设置自动关联
  prop.IsAutoLink = true
  linkField.Apply()
}
main()
```
 
# 145 API文档 / 字段 / 地址字段 / 获取字段值

本页内容

# 获取地址字段类型值 ​

## 说明 ​

获取 地址字段 类型值

## 返回 ​

`object` 结构，结构如下：

### object结构 ​

| key | 值类型 | 说明 |
| --- | --- | --- |
| districts | string[] | 省市区 |
| detail | string | 详细地址 |

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const value = await app.Sheets(1).Views(1).RecordRange(1, "@地址").Value
    console.log(value)
    /**
     * 输出值：
     * {
     *  districts: ["广东省","珠海市","香洲区"],
     *  detail: "前岛环路xxxx号"
     * }
     */
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const value = Application.Sheets(1).Views(1).RecordRange(1, "@地址").Value
    console.log(value)
}
main()
```
 
# 146 API文档 / 字段 / 地址字段 / 设置字段值

本页内容

# 设置地址字段类型值 ​

## 说明 ​

设置 地址字段 类型值

## 返回值 ​

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    // 设置地址字段值
    app.Sheets(1).Views(1).RecordRange(1, "@关联").Value = await Application.DBCellValue({
      districts:["广东省","珠海市","香洲区"],
      detail:"前岛环路xxxx号"
    })
    
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    // 设置关联字段值，参数为传入关联的记录id
    Application.Sheets(1).Views(1).RecordRange(1, "@联系人").Value = Application.DBCellValue({
      districts:["广东省","珠海市","香洲区"],
      detail:"前岛环路xxxx号"
    })
}
main()
```
 
# 147 API文档 / 字段 / 地址字段 / 填写详细地址

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 填写详细地址 ​

AddressField.IsDetailedAddress(属性)

## 说明 ​

可读写

地址字段是否需要填写详细地址

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const Application = instance.Application;
   const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@地址")
   const prop = await fieldDescriptor.Address
   prop.IsDetailedAddress = false
   fieldDescriptor.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@地址")
   const prop = fieldDescriptor.Address
   prop.IsDetailedAddress = false
   fieldDescriptor.Apply()
}
main()
```
 
# 148 API文档 / 字段 / 地址字段 / 预设指定地址

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 预设指定地址 ​

AddressField.IsUsePresetAddress(属性)

## 说明 ​

可读写

地址字段 是否预设指定地址,当设置了IsUsePresetAddress为false后，PresetAddress不生效

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const Application = instance.Application;
   const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@地址")
   const prop = await fieldDescriptor.Address
   prop.IsUsePresetAddress = false
   fieldDescriptor.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@地址")
   const prop = fieldDescriptor.Address
   prop.IsUsePresetAddress = false
   fieldDescriptor.Apply()
}
main()
```
 
# 149 API文档 / 字段 / 地址字段 / 默认值

本页内容

# 默认值 ​

AddressField.PresetAddress(属性)

## 说明 ​

可读写

地址字段的默认值,当设置了IsUsePresetAddress为false后，PresetAddress不生效

## 返回值 ​

Object

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const Application = instance.Application;
   const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@地址")
   const prop = await fieldDescriptor.Address
   prop.IsUsePresetAddress = true
   prop.PresetAddress = {detail:"", districts:["广东省","江门市"]}
   fieldDescriptor.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@地址")
   const prop = fieldDescriptor.Address
   prop.IsUsePresetAddress = true
   prop.PresetAddress = {detail:"", districts:["广东省","江门市"]}
   fieldDescriptor.Apply()
}
main()
```
 
# 150 API文档 / 字段 / 地址字段 / 地址级别数

本页内容

# 地址级别数 ​

AddressField.Level(属性)

## 说明 ​

可读写

地址字段格式包含几级地址，1-5级

| 级数 | 示例 |
| --- | --- |
| 1 | 省 - 详细地址 |
| 2 | 省/市 - 详细地址 |
| 3 | 省/市/区 - 详细地址 |
| 4 | 省/市/区/街道 - 详细地址 |
| 5 | 省/市/区/街道/社区 - 详细地址 |

## 返回值 ​

Number

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const Application = instance.Application;
   const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@地址")
   const prop = await fieldDescriptor.Address
   prop.Level = 5
   fieldDescriptor.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@地址")
   const prop = fieldDescriptor.Address
   prop.Level = 5
   fieldDescriptor.Apply()
}
main()
```
 
# 151 API文档 / 字段 / 公式字段 / 公式的文本表达式

本页内容

# 设置公式的文本表达式 ​

FieldDescriptor.Formula(属性)

## 说明 ​

可读写

公式字段 返回公式的文本表达式

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const Application = instance.Application
  const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@公式")
  const prop = await fieldDescriptor.Formula
  prop.Formula = '=[@日期]+[@日期]'
  fieldDescriptor.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@公式")
  const prop = fieldDescriptor.Formula
  prop.Formula = '=[@日期]+[@日期]'
  fieldDescriptor.Apply()
}
main()
```
 
# 152 API文档 / 字段 / 自动任务字段 / 类型

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 类型 ​

AutomationField.Type(属性)

## 说明 ​

可读写

自动任务字段的类型，有三种类型，参考枚举值 DbAutomationPresetType ，分别为CheckedNotifyContact，UpdatedNotifyContact，DueDateNotifyContact

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const Application = instance.Application;
    const contactField = await Application.FieldDescriptor("Contact","联系人字段")
    await Application.Sheets(1).FieldDescriptors.AddField(contactField)
    const dateField = await Application.FieldDescriptor("Date","日期字段")
    await Application.Sheets(1).FieldDescriptors.AddField(dateField)
    const newField = await Application.FieldDescriptor("Automations","自动任务字段")
    const automation = await newField.Automation
    automation.Type = "DueDateNotifyContact"
    automation.TriggerField = await dateField.Id
    automation.ContactField = await contactField.Id
    automation.ExecuteTime = 3600
    await Application.Sheets(1).FieldDescriptors.AddField(newField)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const contactField = FieldDescriptor("Contact","联系人字段")
    Sheets(1).FieldDescriptors.AddField(contactField)
    const dateField = FieldDescriptor("Date","日期字段")
    Sheets(1).FieldDescriptors.AddField(dateField)
    const newField = FieldDescriptor("Automations","自动任务字段")
    const automation = newField.Automation
    automation.Type = "DueDateNotifyContact"
    automation.TriggerField = dateField.Id
    automation.ContactField = contactField.Id
    automation.ExecuteTime = 3600
    Sheets(1).FieldDescriptors.AddField(newField)
}
main();
```
 
# 153 API文档 / 字段 / 自动任务字段 / 通知联系人

本页内容

# 通知联系人 ​

AutomationField.ContactField(属性)

## 说明 ​

可读写

自动任务字段的通知联系人字段

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const Application = instance.Application;
  const sheet = Application.Sheets(1)
  const field = await sheet.FieldDescriptors(11)
  const automation = await field.Automation
  // 设置 自动化任务的属性
  automation.Type = "DueDateNotifyContact"
  automation.ContactField = await sheet.FieldId("联系人")
  // 更新字段
  field.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const sheet = Application.Sheets(1)
  const field = await sheet.FieldDescriptors(11)
  const automation = await field.Automation
  // 设置 自动化任务的属性
  automation.Type = "DueDateNotifyContact"
  automation.ContactField = sheet.FieldId("联系人")
  // 更新字段
  field.Apply()
}
main();
```
 
# 154 API文档 / 字段 / 自动任务字段 / 触发时间

本页内容

# 触发时间 ​

AutomationField.ExecuteTime(属性)

## 说明 ​

可读写

当自动任务字段的类型为 DueDateNotifyContact，可设置指定触发时间通知联系人 格式为当前日期 + 执行时间, 时间为当天经过的秒数

## 返回值 ​

Number

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const Application = instance.Application;
  const sheet = Application.Sheets(1)
  const field = await sheet.FieldDescriptors(11)
  const automation = await field.Automation
  // 设置 自动化任务的属性
  automation.Type = "DueDateNotifyContact"
  automation.ExecuteTime = 3600
  // 更新字段
  field.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const sheet = Application.Sheets(1)
  const field = await sheet.FieldDescriptors(11)
  const automation = await field.Automation
  // 设置 自动化任务的属性
  automation.Type = "DueDateNotifyContact"
  automation.ExecuteTime = 3600
  // 更新字段
  field.Apply()
}
main();
```
 
# 155 API文档 / 字段 / 级联字段 / 获取字段值

本页内容

# 获取级联字段类型值 ​

## 说明 ​

获取 级联字段 类型值

## 返回 ​

`object` 结构，结构如下

### object结构 ​

| key | 值类型 | 说明 |
| --- | --- | --- |
| districts | string[] | 选项数组 |

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const value = await app.Sheets(1).Views(1).RecordRange(1, "@级联").Value
    console.log(value)
    /**
     * 输出值：
     * {
     *  districts: ["选项1", "选项1-1"], 
     * }
     */
    
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const value = Application.Sheets(1).Views(1).RecordRange(1, "@级联").Value
    console.log(value)
}
main()
```
 
# 156 API文档 / 字段 / 级联字段 / 设置字段值

本页内容

# 设置级联字段类型值 ​

## 说明 ​

设置 级联字段 类型值

## 返回值 ​

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    // 设置级联字段值
    app.Sheets(1).Views(1).RecordRange(2, "@级联").Value = await Application.DBCellValue({
      districts: ['选项1', '选项1-1']
    })
    
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    // 设置级联字段值
    Application.Sheets(1).Views(1).RecordRange(2, "@级联").Value = Application.DBCellValue({
      districts: ['选项1', '选项1-1']
    })
}
main()
```
 
# 157 API文档 / 字段 / 级联字段 / 设置选项

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置选项 ​

CascadeField.AllCascadeOption(属性)

## 说明 ​

可读写

级联字段的选项，最多可以设置4级

## 返回值 ​

[CascadeOptions](/documents/app-integration-dev/guide/dbsheet/Api/CascadeOptions.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const Application = instance.Application;
  const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@级联选项")
  const prop = await fieldDescriptor.Cascade
  const allCascade =  await Application.CascadeOptions();
  const o1 = await allCascade.Add("级联1")
  o1.Children.Add("级联1_1")
  o1.Children.Add("级联1_2")
  const o2 = await allCascade.Add("级联2")
  o2.Children.Add("级联2_1")
  o2.Children.Add("级联2_2")
  const o3 = await allCascade.Add("级联3")
  o3.Children.Add("级联3_1")
  o3.Children.Add("级联3_2")
  const o4 = await allCascade.Add("级联4")
  o4.Children.Add("级联4_1")
  const o4_2 = o4.Children.Add("级联4_2")
  o4_2.Children.Add("级联4_2_1")
  o4_2.Children.Add("级联4_2_2")

  prop.AllCascadeOption = allCascade
  fieldDescriptor.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@级联选项")
  const prop = fieldDescriptor.Cascade
  const allCascade =  CascadeOptions();
  const o1 = allCascade.Add("级联1")
  o1.Children.Add("级联1_1")
  o1.Children.Add("级联1_2")
  const o2 = allCascade.Add("级联2")
  o2.Children.Add("级联2_1")
  o2.Children.Add("级联2_2")
  const o3 = allCascade.Add("级联3")
  o3.Children.Add("级联3_1")
  o3.Children.Add("级联3_2")
  const o4 = allCascade.Add("级联4")
  o4.Children.Add("级联4_1")
  const o4_2 = o4.Children.Add("级联4_2")
  o4_2.Children.Add("级联4_2_1")
  o4_2.Children.Add("级联4_2_2")
  
  prop.AllCascadeOption = allCascade
  fieldDescriptor.Apply()
}
main()
```
 
# 158 API文档 / 字段 / 级联字段 / 显示完整的选择路径

本页内容

# 显示完整的选择路径 ​

CascadeField.IsDisplayAllLevel(属性)

## 说明 ​

可读写

级联字段是否显示完整的选择路径

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const Application = instance.Application
   const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@级联选项")
   const prop = await fieldDescriptor.Cascade
   prop.IsDisplayAllLevel = true
   fieldDescriptor.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@级联选项")
   const prop = fieldDescriptor.Cascade
   prop.IsDisplayAllLevel = true
   fieldDescriptor.Apply()
}
main()
```
 
# 159 API文档 / 字段 / 级联字段 / 选项标题

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 选项标题 ​

FieldDescriptor.Title(属性)

## 说明 ​

可读写

级联选项各级选项的标题，最多可以设置四级标题

## 返回值 ​

Array

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const Application = instance.Application;
  const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@级联选项")
  const prop = await fieldDescriptor.Cascade
  prop.Title = ["第一级标题","第二级标题","第三级标题"]
  fieldDescriptor.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const fieldDescriptor = Application.Sheets(1).FieldDescriptors(2)
  const prop = await fieldDescriptor.Cascade
  prop.Title = ["第一级标题","第二级标题","第三级标题"]
  fieldDescriptor.Apply()
}
main()
```
 
# 160 API文档 / 字段 / 最后修改人 / 最后修改时间字段 / 监听所有字段

本页内容

# 监听所有字段 ​

WatchedField.IsWatchedAll(属性)

## 说明 ​

可读写

最后修改人/最后修改时间，是否监听 所有字段的修改

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const Application = instance.Application
    const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@最后修改时间")
    const prop = await fieldDescriptor.Watch
    prop.IsWatchedAll = true
    fieldDescriptor.Apply()
    console.log(await prop.IsWatchedAll)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@最后修改时间")
    const prop = fieldDescriptor.Watch
    prop.IsWatchedAll = true
    fieldDescriptor.Apply()
}
main()
```
 
# 161 API文档 / 字段 / 最后修改人 / 最后修改时间字段 / 监听某些字段

本页内容

# 监听某些字段 ​

WatchedField.WatchedFields(属性)

## 说明 ​

可读写

最后修改人/最后修改时间，监听某些字段的修改情况, 如果需要监听某些特定的字段，需要将属性 IsWatchedAll 设置为false, 否则这个设置不生效

## 返回值 ​

Array

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const Application = instance.Application
    const fieldDescriptors = await Application.Sheets("数据表").FieldDescriptors
    const fieldDescriptor = await fieldDescriptors.Item("@最后修改时间")
    const prop = await fieldDescriptor.Watch
    const watchId = await fieldDescriptors.Item("@文本").Id
    prop.IsWatchedAll = false
    prop.WatchedFields = [watchId]
    fieldDescriptor.Apply()
    console.log(await prop.WatchedFields)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@最后修改时间")
    const watchId = await Application.Sheets("数据表").FieldDescriptors("@文本").Id
    const prop = fieldDescriptor.Watch
    prop.IsWatchedAll = false
    prop.WatchedFields = [watchId]
    fieldDescriptor.Apply()
}
main()
```
 
# 162 API文档 / 字段 / 监听增加字段的事件

本页内容

- 说明
- 语法
- 参数
- 返回值
- 事件返回数据
- 浏览器环境示例
- 脚本编辑器 示例

# 监听增加字段的事件 ​

FieldDescriptors.OnCreate(方法)

## 说明 ​

为 FieldDescriptors 添加 Create 事件,当添加 FieldDescriptors 时触发。注意在脚本编辑器中使用时，脚本运行结束就会退出运行，这时可能回调无法被正常触发

## 语法 ​

表达式.OnCreate(Callback)

表达式: FieldDescriptors

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Callback | 是 | func | 指定事件发生时的回调方法 ; const eventContext = await FieldDescriptors.OnCreate(()=&gt;{ ... }) |

## 返回值 ​

EventContext

## 事件返回数据 ​

FieldDescriptor

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    let eventContext;
    eventContext = await app.Sheets(1).FieldDescriptors.OnCreate(data => {
        console.log(data);
        // 取消事件监听
        eventContext.Destroy();
    });
    const desc = await app.FieldDescriptor('Rating', '等级字段');
    desc.MaxRating = 2;
    await app.Sheets(1).FieldDescriptors.AddField(desc, 1);
    //这里会执行OnCreate的回调
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    let eventContext;
    eventContext = Application.Sheets(1).FieldDescriptors.OnCreate(data => {
        console.log(data);
        // 取消事件监听
        eventContext.Destroy();
    });
    const desc = Application.FieldDescriptor('Rating', '等级字段');
    desc.MaxRating = 2;
    Application.Sheets(1).FieldDescriptors.AddField(desc, 1);
    //这里会执行OnCreate的回调
}
main();
```
 
# 163 API文档 / 字段 / 监听删除字段的事件

本页内容

# 监听删除字段的事件 ​

FieldDescriptor.OnDelete(方法)

## 说明 ​

为 FieldDescriptor 添加 Delete 事件,当删除 FieldDescriptor 时触发。注意在脚本编辑器中使用时，脚本运行结束就会退出运行，这时可能回调无法被正常触发

## 语法 ​

表达式.OnDelete(Callback)

表达式: FieldDescriptor

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Callback | 是 | func | 指定事件发生时的回调方法 ; const eventContext = await FieldDescriptor.OnDelete(()=&gt;{ ... }) |

## 返回值 ​

EventContext

## 事件返回数据 ​

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| sheetId | Number | 表的 Id |
| fieldId | String | 字段的 Id |
| fieldIds | Array | 字段集合的 Ids |

## 事件返回数据示例 ​

```javascript
{
    fieldId: "C"
    fieldIds: ['C']
    sheetId: 1
}
```

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const field = await app.Sheets(1).FieldDescriptors(2);
    let eventContext;
    eventContext = await field.OnDelete(data => {
        console.log(data);
        // 取消事件监听
        eventContext.Destroy();
    });
    await field.Delete();
    //这里会执行OnDelete的回调
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const field = Application.Sheets(1).FieldDescriptors(2);
    let eventContext;
    eventContext = field.OnDelete(data => {
        console.log(data);
        // 取消事件监听
        eventContext.Destroy();
    });
    field.Delete();
    field.Delete();
    //这里会执行OnDelete的回调
}
main();
```
 
# 164 API文档 / 字段 / 监听修改字段的事件

本页内容

# 监听修改字段的事件 ​

FieldDescriptor.OnUpdate(方法)

## 说明 ​

为 FieldDescriptor 添加 Update 事件,当更新 FieldDescriptor 时触发。注意在脚本编辑器中使用时，脚本运行结束就会退出运行，这时可能回调无法被正常触发

## 语法 ​

表达式.OnUpdate(Callback)

表达式: FieldDescriptor

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Callback | 是 | func | 指定事件发生时的回调方法 ; const eventContext = await FieldDescriptor.OnUpdate(()=&gt;{ ... }) |

## 返回值 ​

EventContext

## 事件返回数据 ​

FieldDescriptor

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const fieldId = app.Sheets(1).FieldId('修改前字段名');
    const fieldDescriptor = app.Sheets(1).FieldDescriptors(fieldId);
    let eventContext
    eventContext = await fieldDescriptor.OnUpdate(data => {
        console.log(data);
        // 取消事件监听
        eventContext.Destroy()
    });

    fieldDescriptor.Name = '修改字段名';
    fieldDescriptor.Apply();
    //这里会执行OnUpdate的回调
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const fieldDescriptor = Application.Sheets(1).FieldDescriptors(1);
    let eventContext
    eventContext = fieldDescriptor.OnUpdate(data => {
        console.log(data);
        // 取消事件监听
        eventContext.Destroy()
    });
    fieldDescriptor.Name = '修改字段名';
    fieldDescriptor.Apply();
    //这里会执行OnUpdate的回调
}
main();
```
 
# 165 API文档 / 排序 / 设置排序升序属性

本页内容

# 设置排序升序属性 ​

Sort.IsAscending(属性)

## 说明 ​

可读写

设置排序升序或者降序

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
// 获取IsAscending属性
async function example() {
    await instance.ready();
    const app = instance.Application;
    const sorts = await app.Sheets(1).Views(1).Sorts;
    const isAscending = sorts(1).IsAscending;
}

// 设置IsAscending属性
async function example() {
    await instance.ready();
    const app = instance.Application;
    app.Sheets(1).Views(1).Sorts(1).IsAscending = false;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const sorts = Application.Sheets(1).Views(1).Sorts;
    const isAscending = sorts(1).IsAscending;
    isAscending = false;
}
main()
```
 
# 166 API文档 / 排序 / 添加排序条件

本页内容

# 添加排序条件 ​

Sorts.Add(方法)

## 说明 ​

添加排序

## 语法 ​

表达式.Add(Field,IsAscending)

表达式: Sorts

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Field | 是 | number/string | 新增排序字段索引/新增排序字段 ID/新增排序字段名(名称要以@字符作为开始) |
| IsAscending | 否 | boolean | 是否为升序 |

## 返回值 ​

Sort

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const sorts = await app.Sheets(1).Views(1).Sorts;
    const res = sorts.Add(1);
    //const res = sorts.Add('B');
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const sorts = Application.Sheets(1).Views(1).Sorts;
    const res = sorts.Add(1);
    //const res = sorts.Add('B');
}
main()
```
 
# 167 API文档 / 排序 / 移动排序条件

本页内容

# 移动排序条件 ​

Sorts.ChangeOrder(方法)

## 说明 ​

移动排序条件(设置排序优先级)

## 语法 ​

表达式.ChangeOrder(FromField, BeforeField, AfterField)

表达式: Sorts

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| FromField | 是 | [string] | 要移动的排序字段的字段ID/要移动的排序字段的字段名称(名称要以@字符作为开始) |
| BeforeField | 否 | [string] | 目标位置前的排序字段ID/目标位置前的排序字段名称(名称要以@字符作为开始) |
| AfterField | 否 | [string] | 目标位置后的排序字段ID/目标位置后的排序字段名称(名称要以@字符作为开始) |

FromField、BeforeField和AfterField必须都是已设置的排序条件字段，BeforeField和AfterField至少需要传入一个，如果BeforeField和AfterField同时存在以BeforeField作为应用参数

比如表格视图中已设置的排序条件在排序面板中从上到下依次为【公式，日期，名称，数量】 现在想将名称这条排序条件移动到日期的前面，结果变为【公式，名称，日期，数量】，就可以用以下方式实现

```javascript
await WPSOpenApi.Application.Sheets(1).Views(1).Sorts.ChangeOrder('@名称', '@日期')
```

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const sorts = await app.Sheets(1).Views(1).Sorts;
    // 将公式排序条件移动到日期排序条件的后面
    const res = await sorts.ChangeOrder('@公式', undefined, '@日期');
    if (res.Code === 0) {
        console.log("设置排序优先级成功")
    } else {
        console.error("设置排序优先级失败" + res.Message)
    }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const sorts = Application.Sheets(1).Views(1).Sorts;
    // 将公式排序条件移动到日期排序条件的后面
    const res = sorts.ChangeOrder('@公式', undefined, '@日期');
    if (res.Code === 0) {
        console.log("设置排序优先级成功")
    } else {
        console.error("设置排序优先级失败" + res.Message)
    }
}
main()
```
 
# 168 API文档 / 筛选 / 设置筛选条件

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置筛选条件 ​

Filter.Criteria(属性)

## 说明 ​

可读写

设置或获取 单条筛选记录的条件

![alt text](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/Filter_Criteria.Vsk-X2TT.png)

## 返回值 ​

[Criteria](/documents/app-integration-dev/guide/dbsheet/Api/Criteria.html)

## 浏览器环境示例 ​

javascript

```javascript
// 获取单条筛选记录的条件
async function example() {
    await instance.ready();
    const app = instance.Application;
    // 获取筛选
    const filters = await app.Sheets(1).Views(1).Filters;
    // 获取第一个条件
    const criteria = await filters.Item(1).Criteria;
    console.log(criteria.Op) // "Equals"
    console.log(criteria.Field) // 1
    console.log(criteria.Values[0]) // {type: 'Text', value: '1'}
}

// 设置单条筛选记录的条件
async function example() {
    await instance.ready();
    const app = instance.Application;
    // 构造筛选数据
    const Criteria = app.Criteria(1, "Equals", ["1"])
    // 设置筛选数据
    app.Sheets(1).Views(1).Filters.Item(1).Criteria = Criteria;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const criteria = Application.Sheets(1).Views(1).Filters.Item(1).Criteria;
    criteria = Criteria(1, "Equals", ["1"]);
}
main();
```
 
# 169 API文档 / 筛选 / 添加筛选条件

本页内容

# 添加筛选条件 ​

Filters.Add(方法)

## 说明 ​

添加筛选条件

![alt text](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/Filter_Criteria.Vsk-X2TT.png)

## 语法 ​

表达式.Add(Criteria)

表达式:Filters

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Criteria | 是 | Object | 筛选条件 |

## 返回值 ​

Filter

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const filters = await app.Sheets(1).Views(1).Filters;
    const criteria = app.Criteria(1, 'Equals', ['1'])
    const filter = await filters.Add(criteria);
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const filters = Application.Sheets(1).Views(1).Filters;
    const criteria = Application.Criteria(1, 'Equals', ['1'])
    const filter = filters.Add(criteria);
}
main()
```
 
# 170 API文档 / 筛选 / 删除筛选条件

本页内容

# 删除筛选条件 ​

Filter.Delete(方法)

## 说明 ​

删除单条筛选记录

## 语法 ​

表达式.Delete()

表达式:Filter

## 参数 ​

无参数

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const res = await app.Sheets(1).Views(1).Filters(1).Delete();
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const res = Application.Sheets(1).Views(1).Filters(1).Delete();
}
main();
```
 
# 171 API文档 / 分组 / 添加分组

本页内容

# 添加分组 ​

Groups.Add(方法)

## 说明 ​

添加分组

## 语法 ​

表达式.Add(Field, IsAscending)

表达式:Groups

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Field | 是 | number/string | 新增分组字段索引/新增分组字段ID/新增分组字段名 |
| IsAscending | 否 | boolean | 是否是升序 |

## 返回值 ​

Group

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const res = await app.Sheets(1).Views(1).Groups.Add(1);
    // const res = await app.Sheets(1).Views(1).Groups.Add("@数量");
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const res = Application.Sheets(1).Views(1).Groups.Add(1);
    // const res = Application.Sheets(1).Views(1).Groups.Add("@数量");
}
main();
```
 
# 172 API文档 / 分组 / 删除分组

本页内容

# 删除分组 ​

Group.Delete(方法)

## 说明 ​

删除分组条件

## 语法 ​

表达式.Delete()

表达式:Group

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |

## 返回值 ​

ApiResult

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const res = await app.Sheets(1).Views(1).Groups(1).Delete();
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const res = Application.Sheets(1).Views(1).Groups(1).Delete();
}
main()
```
 
# 173 API文档 / 分组 / 展开分组

本页内容

# 展开分组 ​

Groups.UnFoldAll(方法)

## 说明 ​

展开分组

## 语法 ​

表达式.UnFoldAll()

表达式:Groups

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const res = await app.Sheets(1).Views(1).Groups.UnFoldAll();
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const res = Application.Sheets(1).Views(1).Groups.UnFoldAll();
}
main()
```
 
# 174 API文档 / 分组 / 分组折叠

本页内容

# 折叠分组 ​

Groups.FoldAll(方法)

## 说明 ​

折叠分组

## 语法 ​

表达式.FoldAll()

表达式:Groups

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const res = await app.Sheets(1).Views(1).Groups.FoldAll();
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const res = Application.Sheets(1).Views(1).Groups.FoldAll();
}
main()
```
 
# 175 API文档 / 评论 / 插入

本页内容

# 插入新的评论 ​

RecordComment.Add(方法)

## 说明 ​

插入新的评论

## 语法 ​

表达式.Add(Text, TextLinkRuns)

表达式:RecordComment

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Text | 是 | String | 评论的文本，插入到评论的最前方 |
| TextLinkRuns | 否 | Array | 文本的特殊节点属性 |

## 返回值 ​

DbComment

## jsApi 示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const app = instance.Application;
   const recordComment = await app.ActiveView.RecordComments.Item(1)
   recordComment.Add("Hello World")
   const count = await recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = await recordComment.Item(i)
    console.log(await comment.Text)
   }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const recordComment = ActiveView.RecordComments.Item(1)
   recordComment.Add("Hello World")
   const count = recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = recordComment.Item(i)
    console.log(comment.Text)
   }
}
main()
```
 
# 176 API文档 / 评论 / 删除

本页内容

# 删除评论 ​

RecordComment.Delete(方法)

## 说明 ​

删除评论

## 语法 ​

表达式.Delete()

表达式:RecordComment

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Index | 是 | Number/String | 删除记录的索引 |

## 返回值 ​

ApiResult

## jsApi 示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const app = instance.Application;
   const recordComment = await app.ActiveView.RecordComments.Item(1)
   const count = await recordComment.Count
   for (let i = 1; i <= count; i++) { 
    await recordComment.Delete(i)
   }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const recordComment = ActiveView.RecordComments(1)
   const count = recordComment.Count
   for (let i = 1; i <= count; i++) { 
    await recordComment.Delete(i)
   }
}
main()
```
 
# 177 API文档 / 评论 / 监听插入评论

本页内容

# 监听插入评论 ​

RecordComments.OnCreate(方法)

## 说明 ​

为 RecordComments 添加 OnCreate 事件,当创建 评论 时触发。注意在脚本编辑器中使用时，脚本运行结束就会退出运行，这时可能回调无法被正常触发 这个方法只能监听视图的事件， 如果在浏览器环境需要全局监听也可以使用

javascript

```javascript
jssdk.on("OnBroadcast", async (res) => {
    const data = res.Data
    if (data.type == "DB_COMMENT_UPDATE") { // 收到文档评论更新消息
        if (data.shouldNotLocalUpdate) {
            // 本地更新评论信息
            console.log("收到广播消息：", data)
            const info = data.info
            const {sheetStId, commentId, recordId, action} = info
            if (action == "Add") {
                // 新增评论
                const addText = await jssdk.Application.Sheets.ItemById(sheetStId).ActiveView.RecordComments(recordId).Item(commentId).Text
                console.log("新增评论：", addText)
            } else if (action == "Delete") {
                // 删除评论
                console.log("删除评论：", info)
            }
        }
    }
})
```

可以通过 action 来判断是哪个事件触发的

## 语法 ​

表达式.OnDelete(Callback)

表达式: RecordComments

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Callback | 是 | func | 指定事件发生时的回调方法 ; const eventContext = await RecordComments.OnCreate(()=&gt;{ ... }) |

## 返回值 ​

EventContext

## 事件返回数据 ​

DbComment

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    let eventContext;
    eventContext = await app.ActiveView.RecordComments.OnCreate(async (data)=> {
        const info = data.info
        const {sheetStId, commentId, recordId, action} = info
        if (action == "Add") {
            // 新增评论
            const addText = await app.Sheets.ItemById(sheetStId).ActiveView.RecordComments(recordId).Item(commentId).Text
            console.log("新增评论：", addText)
        } else if (action == "Delete") {
            // 删除评论
            console.log("删除评论：", info)
        }
    })

    // 移除监听
    // eventContext.Destroy();
}
```
 
# 178 API文档 / 评论 / 监听删除评论

本页内容

# 监听删除评论 ​

RecordComments.OnDelete(方法)

## 说明 ​

为 RecordComments 添加 Delete 事件,当删除 评论 时触发。注意在脚本编辑器中使用时，脚本运行结束就会退出运行，这时可能回调无法被正常触发 这个方法只能监听视图的事件， 如果在浏览器环境需要全局监听也可以使用

jssdk.on("OnBroadcast", (res)=&gt;console.error("##", res))

回调的消息数据包含的内容跟事件返回数据是一致的， 可以通过 action 来判断是哪个事件触发的

## 语法 ​

表达式.OnDelete(Callback)

表达式: RecordComments

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Callback | 是 | func | 指定事件发生时的回调方法 ; const eventContext = await RecordComments.OnDelete(()=&gt;{ ... }) |

## 返回值 ​

EventContext

## 事件返回数据 ​

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| commentId | String | 评论ID |
| recordId | String | 记录ID |
| sheetStId | Number | 表ID |

## 事件返回数据示例 ​

```javascript
{"recordId":"Bk","sheetStId":1,"commentId":"e66e42020baa4d5455da5d2043c631a5","action":"Delete"}
```

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    let eventContext;
    eventContext = await app.ActiveView.RecordComments.OnDelete((data)=>console.error(JSON.stringify(data)))
    await app.ActiveView.RecordComments(1).Item(1).Delete()

    // 移除监听
    // eventContext.Destroy();
}
```
 
# 179 API文档 / 同步数据 / 刷新数据

本页内容

# 刷新数据 ​

SyncDBSheet.RefreshSyncSheet(方法)

JSSDK v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

刷新数据

## 语法 ​

表达式.RefreshSyncSheet()

表达式:SyncDBSheet

## 参数 ​

无参数

## 返回值 ​

Boolean

## jsAPI示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 // 切到某个同步表，获取该同步表的实例对象
 const syncSheet = await app.ActiveSheet
 await syncSheet.RefreshSyncSheet()
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 // 切到某个同步表，获取该同步表的实例对象
 const syncSheet = Application.ActiveSheet
 syncSheet.RefreshSyncSheet()
}
main()
```
 
# 180 API文档 / 同步数据 / 解除同步关系

本页内容

# 解除同步关系 ​

SyncDBSheet.RemoveSheetSyncLink(方法)

JSSDK v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

解除同步关系

## 语法 ​

表达式.RemoveSheetSyncLink()

表达式:SyncDBSheet

## 参数 ​

无参数

## 返回值 ​

Boolean

## jsAPI示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 // 切到某个同步表，获取该同步表的实例对象
 const syncSheet = await app.ActiveSheet
 await syncSheet.RemoveSheetSyncLink()
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 // 切到某个同步表，获取该同步表的实例对象
 const syncSheet = Application.ActiveSheet
 syncSheet.RemoveSheetSyncLink()
}
main()
```
 
# 181 API文档 / 同步数据 / 创建同步表

本页内容

# 创建同步表 ​

DataSource.CreateSyncDBSheets(方法)

## 说明 ​

创建同步表

## 语法 ​

表达式.CreateSyncDBSheets(FileId, OfficeType, SheetIds) 表达式:DataSource

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| FileId | 是 | string | 在线表格文件的 fileId |
| OfficeType | 是 | 'd' | 'k' | 在线表格文件的类型：d 代表多维表格，k 代表智能表格 |
| SheetIds | 是 | number[] | 在线表格文件的 stId 数组 |

## 返回值 ​

SyncDBSheet[]

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const syncSheets = await app.DataSource.CreateSyncDBSheets('100127684526', 'd', [1])
    console.log(syncSheets)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const syncSheets = Application.DataSource.CreateSyncDBSheets('100127684526', 'd', [1])
  console.log(syncSheets)
}
main()
```
 
# 182 API文档 / 同步数据 / 创建合并表

本页内容

# 创建合并表 ​

DataSource.CreateSummarySheet(方法)

JSSDK v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

将数据源表合并成一个合并表。

## 语法 ​

表达式.CreateSummarySheet(SummarySourceConfigs)

表达式:DataSource

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| SummarySourceConfigs | 是 | SummarySourceConfigs | 数据源配置对象 |

## 返回值 ​

SummarySheet

## jsAPI示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    // 通过SummarySourceConfigs对象构造符合规范的数据源配置对象，来创建合并表
    const configs = await app.SummarySourceConfigs;
    await configs.Add('100138603654', [1])
    const sheet = await app.DataSource.CreateSummarySheet(configs)
    console.log(sheet)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  // 通过SummarySourceConfigs对象构造符合规范的数据源配置对象，来创建合并表
  const configs = Application.SummarySourceConfigs;
  configs.Add('100138603654', [1])
  const sheet = Application.DataSource.CreateSummarySheet(configs)
  console.log(sheet)
}
main()
```
 
# 183 API文档 / 同步数据 / 合并表-添加数据源

本页内容

# 添加数据源 ​

SummarySourceConfigs.Add(方法)

JSSDK v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

合并表配置对象中添加数据源（只是编辑的本地对象，需要调Apply方法才能更新到云上）

## 语法 ​

表达式.Add(File,Sheets)

表达式:SummarySourceConfigs

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| File | 是 | String | 文件id/文件url |
| Sheets | 是 | Number/String[] | 数据表数组，支持两种形式，表名和索引（从1开始） |

## 返回值 ​

SummarySourceConfig

## jsAPI示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const configs = await app.ActiveSheet.SourceConfigs
 const config = await configs.Add("100136699885", [1, 2])
 // 返回配置文件的id
 console.log(await config.FileId)
 // 返回文件下选中的数据表id数组
 console.log(await config.SheetIds)
 // 将改动后的配置更新到云上
 await configs.Apply()
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 const configs = Application.ActiveSheet.SourceConfigs
 const config = configs.Add("100136699885", [1, 2])
 // 返回配置文件的id
 console.log(config.FileId)
 // 返回文件下选中的数据表id数组
 console.log(config.SheetIds)
 // 将改动后的配置更新到云上
 configs.Apply()
}
main()
```
 
# 184 API文档 / 同步数据 / 合并表-删除数据源

本页内容

# 删除数据源 ​

SummarySourceConfigs.Delete(方法)

JSSDK v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

合并表配置对象中删除数据源（只是编辑的本地对象，需要调Apply方法才能更新到云上）

## 语法 ​

表达式.Delete(Index)

表达式:SummarySourceConfigs

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Index | 是 | Number/String | 支持索引、文件id和文件url |

## 返回值 ​

Boolean

## jsAPI示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const configs = await app.ActiveSheet.SourceConfigs
 await configs.Delete(1)
 // 将改动后的配置更新到云上
 await configs.Apply()
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 const configs = Application.ActiveSheet.SourceConfigs
 configs.Delete(1)
 // 将改动后的配置更新到云上
 configs.Apply()
}
main()
```
 
# 185 API文档 / 公告 / 公告-展开 / 收起

本页内容

# 展开/收起 ​

NoticeBar.Visible(属性)

## 说明 ​

可读写

设置公告栏可见性

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function getVisible() {
    await instance.ready();
    const app = instance.Application;
    const noticeBar = await app.Window.NoticeBar
    const visible =  await noticeBar.Visible     // 为true公告栏隐藏，为false公告栏显示
}
async function setVisible() {
    await instance.ready();
    const app = instance.Application;
    const noticeBar = await app.Window.NoticeBar
    noticeBar.Visible = false     // 隐藏公告栏
    noticeBar.Visible = true      // 显示公告栏
}
```
 
# 186 API文档 / 公告 / 窗口公告栏

本页内容

# 窗口公告栏 ​

Window.NoticeBar(属性)

## 说明 ​

返回窗口公告栏对象信息

## 返回值 ​

NoticeBar

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const noticeBar = await app.Window.NoticeBar
}
```
 
# 187 API文档 / 其他 / 设置导航栏可见性

本页内容

# 设置导航栏可见性 ​

Navigator.Visible(属性)

## 说明 ​

可读写

设置导航栏隐藏或显示

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function getVisible() {
    await instance.ready();
    const app = instance.Application;
    const navigator = await app.Window.Navigator
    const visible =  await navigator.Visible     // 为true导航栏隐藏，为false导航栏显示
}
async function setVisible() {
    await instance.ready();
    const app = instance.Application;
    const navigator = await app.Window.Navigator
    navigator.Visible = false     // 隐藏导航栏
    navigator.Visible = true      // 显示导航栏
}
```
 
# 188 API文档 / 其他 / 代理界面元素

本页内容

# 代理界面元素 ​

Window.BailHook(方法)

## 说明 ​

对特定的界面元素进行代理，代理方法返回true时，则不显示原界面

## 语法 ​

表达式.BailHook(CmbId)

表达式: Window

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| CmbId | 是 | string | 界面元素ID，目前支持 `RecordInfo`(记录详情卡片) |

## 返回值 ​

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;

    const hook = await app.Window.BailHook("RecordInfo")
    hook.InvokeSingle((params)=>{
        console.log(params) // 移动端返回参数 {recordId: 'Jp', activeFieldId: 'E'}
                            // pc端返回参数 {recordId: 'Jp', isShowComment: false}
                            // 移动端和PC端 用到的参数都是 recordId，其它参数是界面区别
        // 可以通过 await WPSOpenApi.Application.ActiveView.RecordRange(params.recordId).Value 读取记录的数据
        const count = await Application.ActiveView.RecordRange.Count
        const record = Application.ActiveView.RecordRange(params.recordId)
        const values = await record.Value
        const index = await record.Index // 注意 Index base 1, 可能有多条记录，返回 [index]
        const prevRecord = Application.ActiveView.RecordRange(index.map(_=> _ - 1))
        const nextRecord = Application.ActiveView.RecordRange(index.map(_=> _ + 1))
        // 在这里实现自定义的界面逻辑，替换掉原来的界面
        record.Select()
        return true // return true 会不弹出原界面
    })
}
```
 
# 189 API文档 / 其他 / 设置经典布局

本页内容

# 设置经典布局 ​

Window.SetLayout(方法)

## 说明 ​

设置该视图是否为经典布局

## 语法 ​

表达式.SetLayout(isClassic)

表达式: Window

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| isClassic | 是 | Boolean | 是否为经典布局 |

## 返回值 ​

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    await app.Window.SetLayout(true);
    await app.Window.SetLayout(false);
}
```
 
# 190 API文档 / 其他 / 窗口导航栏

本页内容

# 窗口导航栏 ​

Window.Navigator(属性)

## 说明 ​

返回窗口导航栏对象信息

## 返回值 ​

Navigator

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const navigator = await app.Window.Navigator
}
```
 
# 191 API文档 / API / Sheet / Sheet对象

本页内容

# Sheet 数据表操作 ​

## 说明 ​

- sheet: 泛指数据表（dbSheet）、仪表盘（dashboardSheet）、说明页面（fpSheet）

## 方法 ​

| 方法 | 说明 |
| --- | --- |
| [AddDescription](/documents/app-integration-dev/guide/dbsheet/Api/Sheet_AddDescription.html) | 为当前数据表添加说明 |
| [Copy](/documents/app-integration-dev/guide/dbsheet/Api/Sheet_Copy.html) | 创建副本 |
| [CreateFields](/documents/app-integration-dev/guide/dbsheet/Api/Sheet_CreateFields.html) | 创建字段 |
| [Delete](/documents/app-integration-dev/guide/dbsheet/Api/Sheet_Delete.html) | 删除当前数据表 |
| [DeleteFields](/documents/app-integration-dev/guide/dbsheet/Api/Sheet_DeleteFields.html) | 删除指定字段信息 |
| [FieldId](/documents/app-integration-dev/guide/dbsheet/Api/Sheet_FieldId.html) | 通过字段名称来获取字段Id |
| [GetFields](/documents/app-integration-dev/guide/dbsheet/Api/Sheet_GetFields.html) | 获取该数据表的所有字段信息 |
| [Share](/documents/app-integration-dev/guide/dbsheet/Api/Sheet_Share.html) | sheet分享 |
| [UpdateFields](/documents/app-integration-dev/guide/dbsheet/Api/Sheet_UpdateFields.html) | 更新字段 |
| [AppendFromLocal](/documents/app-integration-dev/guide/dbsheet/Api/Sheet_AppendFromLocal.html) | 将本地表格文件数据追加导入到当前数据表 |
| [AppendFromCloud](/documents/app-integration-dev/guide/dbsheet/Api/Sheet_AppendFromCloud.html) | 将在线表格文件数据追加导入到当前数据表 |

## 属性 ​

| 属性 | 说明 | 读写说明 |
| --- | --- | --- |
| [FieldDescriptors](/documents/app-integration-dev/guide/dbsheet/Api/Sheet_FieldDescriptors.html) | 字段描述集合 | 可读写 |
| [Icon](/documents/app-integration-dev/guide/dbsheet/Api/Sheet_Icon.html) | 图标 | 可读写 |
| [Id](/documents/app-integration-dev/guide/dbsheet/Api/Sheet_Id.html) | Id | 可读 |
| [Name](/documents/app-integration-dev/guide/dbsheet/Api/Sheet_Name.html) | 名称 | 可读写 |
| [RecordRange](/documents/app-integration-dev/guide/dbsheet/Api/Sheet_RecordRange.html) | 所有记录 | 可读 |
| [Views](/documents/app-integration-dev/guide/dbsheet/Api/Sheet_Views.html) | 视图集合 | 可读 |

## 事件 ​

| 事件 | 说明 |
| --- | --- |
| [OnCreateRecord](/documents/app-integration-dev/guide/dbsheet/Api/Sheet_OnCreateRecord.html) | 新增 Record 时触发 |
| [OnDelete](/documents/app-integration-dev/guide/dbsheet/Api/Sheet_OnDelete.html) | 删除 Sheet 时触发 |
| [OnRename](/documents/app-integration-dev/guide/dbsheet/Api/Sheet_OnRename.html) | 修改 Name 时触发 |

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const sheet = await app.Sheets(1);
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const sheet = await app.Sheets(1);
}
main()
```
 
# 192 API文档 / API / Sheet / FieldDescriptors

本页内容

# 字段描述集合 ​

## 说明 ​

可读写

该表的字段描述集合

## 返回值 ​

[FieldDescriptors](/documents/app-integration-dev/guide/dbsheet/Api/FieldDescriptors.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const fieldDescriptors = app.Sheet(1).FieldDescriptors;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const fieldDescriptors = Application.Sheet(1).FieldDescriptors;
}
main()
```
 
# 193 API文档 / API / Sheet / Icon

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置图标 ​

## 说明 ​

可读写

设置图标、返回当前数据表的图标

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const sheet = app.Sheets(1);
    // read
    const sheetIcon = await sheet.Icon;
    // write
    sheet.Icon = '📚';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const sheet = Application.Sheets(1);
    // read
    const sheetIcon =  sheet.Icon;
    // write
    sheet.Icon = '📚';
}
main()
```
 
# 194 API文档 / API / Sheet / Id

本页内容

# 返回数据表Id ​

## 说明 ​

可读 返回当前数据表的 Id

## 返回值 ​

Number

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const sheetId = await app.Sheets(1).Id;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const sheetId = Application.Sheets(1).Id;
}
main()
```
 
# 195 API文档 / API / Sheet / Name

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 重命名 ​

## 说明 ​

可读写

重命名、返回当前数据表的名称

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const sheet = app.Sheets(1);
    // read
    const sheetName = await sheet.Name;
    // write
    sheet.Name = 'newSheetName';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const sheet = Application.Sheets(1);
    // read
    const sheetName = sheet.Name;
    // write
    sheet.Name = 'newSheetName';
}
main()
```
 
# 196 API文档 / API / Sheet / RecordRange

本页内容

# 数据表的所有记录 ​

## 说明 ​

可读 该数据表的所有记录

## 返回值 ​

RecordRange

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const recordRange = await app.Sheet(1).RecordRange;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const recordRange = Application.Sheet(1).RecordRange;
}
main()
```
 
# 197 API文档 / API / Sheet / Views

本页内容

# 数据表的视图集合 ​

## 说明 ​

可读 返回当前数据表的视图集合

## 返回值 ​

Views

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const views = await app.Sheet(1).Views;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const views = Application.Sheet(1).Views;
}
main()
```
 
# 198 API文档 / API / Sheet / AddDescription

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 添加说明 ​

## 说明 ​

为当前数据表添加说明

## 语法 ​

表达式.AddDescription(Value)

表达式: Sheet

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Value | 是 | string | 待添加的说明文案 |

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    await app.Sheets(1).AddDescription('hello');
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  Application.Sheets(1).AddDescription('hello');
}
main()
```
 
# 199 API文档 / API / Sheet / AppendFromCloud

本页内容

# 将在线表格文件数据追加导入到当前数据表 ​

JSSDK v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

将在线表格文件数据追加导入到当前数据表

## 语法 ​

表达式.AppendFromCloud(FileId)

表达式: Sheet

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| FileId | 是 | string | 在线表格文件的 fileId |

## 返回值 ​

Boolean

## jsAPI示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const curSheet = app.Sheets.Item(1)
    try {
      const isAppended = await curSheet.AppendFromCloud(fileId)
      if (isAppended) {
        console.log('“追加数据到数据表成功')
      }
    } catch (error) {
      console.log('“追加数据到数据表失败', error.message)
    }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  try {
    const isAppended = Application.Sheets.Item(1).AppendFromCloud(fileId)
    if (isAppended) {
      console.log('“追加数据到数据表成功')
    }
  } catch (error) {
    console.log('“追加数据到数据表失败', error.message)
  }
}
main()
```
 
# 200 API文档 / API / Sheet / AppendFromLocal

本页内容

# 将本地表格文件数据追加导入到当前数据表 ​

JSSDK v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

将本地表格文件数据追加导入到当前数据表

## 语法 ​

表达式.AppendFromLocal(File)

表达式: Sheet

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| File | 是 | File | 本地表格文件对象 |

## 返回值 ​

Boolean

## jsAPI示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const curSheet = app.Sheets.Item(1)
    
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];
    if (!file) {
        console.error('没有选择文件');
        return;
    }
    
    try {
      const isAppended = await curSheet.AppendFromLocal(file)
      if (isAppended) {
        console.log('“追加数据到数据表成功')
      }
    } catch (error) {
      console.log('“追加数据到数据表失败', error.message)
    }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  try {
    const isAppended = Application.Sheets.Item(1).AppendFromLocal(file)
    if (isAppended) {
      console.log('“追加数据到数据表成功')
    }
  } catch (error) {
    console.log('“追加数据到数据表失败', error.message)
  }
}
main()
```
 
# 201 API文档 / API / Sheet / Copy

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 创建副本 ​

## 说明 ​

为当前数据表创建副本

## 语法 ​

表达式.Copy(Value)

表达式: Sheet

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Value | 否 | boolean | 创建副本的方式，默认为 false。传 true 复制全部内容；传 false 仅复制空表和视图；复制的sheet为仪表盘时，此参数不传 |

## 返回值 ​

Sheet

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    // 复制全部内容
    await app.Sheets(1).Copy(true);
    // 仅复制空表和视图
    await app.Sheets(1).Copy(false);
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    // 复制全部内容
    Application.Sheets(1).Copy(true);
    // 仅复制空表和视图
    Application.Sheets(1).Copy(false);
}
main()
```
 
# 202 API文档 / API / Sheet / CreateFields

本页内容

# 创建字段 ​

## 说明 ​

为当前数据表创建字段

## 语法 ​

表达式.CreateFields(Fields)

表达式: Sheet

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Fields | 是 | Array | 创建的字段数组 |

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    await app.ActiveSheet.CreateFields([{name:'等级',type:'Rating',max:5},{name:'富文本',type:'MultiLineText'}])
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    Application.ActiveSheet.CreateFields([{name:'等级',type:'Rating',max:5},{name:'富文本',type:'MultiLineText'}])
}
main()
```
 
# 203 API文档 / API / Sheet / Delete

本页内容

# 删除数据表 ​

## 说明 ​

删除当前数据表

## 语法 ​

表达式.Delete(HideTip)

表达式: Sheet

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| HideTip | 否 | boolean | 是否需要隐藏删除确认框 |

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    await app.Sheets(1).Delete(true);
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  Application.Sheets(1).Delete(true);
}
main()
```
 
# 204 API文档 / API / Sheet / DeleteFields

本页内容

# 删除指定字段信息 ​

## 说明 ​

删除该数据表的指定 ids 的字段信息

## 语法 ​

表达式.DeleteFields(FiledIds)

表达式: Sheet

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| FiledIds | 是 | Array | 要删除的字段的 ids(字符串数组) |

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    await app.ActiveSheet.DeleteFields(['B','C']);
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    Application.ActiveSheet.DeleteFields(['B','C']);
}
main();
```
 
# 205 API文档 / API / Sheet / FieldId

本页内容

# 通过字段名称来获取字段Id ​

## 说明 ​

通过字段名称来获取字段Id

## 语法 ​

表达式: FieldId()

表达式: Sheet

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Name | 是 | String | 字段名称 |

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const fieldId = await app.Sheets(1).FieldId('字段名');
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const fieldId = Application.Sheets(1).FieldId('字段名'); 
}
main()
```
 
# 206 API文档 / API / Sheet / GetFields

本页内容

# 获取该数据表的所有字段信息 ​

## 说明 ​

获取该数据表的所有字段信息

## 语法 ​

表达式.GetFields()

表达式: Sheet

## 参数 ​

无参数

## 返回值 ​

[Fields](/documents/app-integration-dev/guide/dbsheet/Api/Fields.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const fields = app.ActiveSheet.GetFields();
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const fields = Application.ActiveSheet.GetFields(); 
}
main()
```
 
# 207 API文档 / API / Sheet / Share

本页内容

# sheet分享 ​

## 说明 ​

- 当前方法暂时仅支持仪表盘sheet分享

## 语法 ​

表达式.Share()

表达式: Sheet

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    // 分享仪表盘
    await app.Sheets(1).Share();
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    // 分享仪表盘
    Application.Sheets(1).Share(); 
}
main()
```
 
# 208 API文档 / API / Sheet / UpdateFields

本页内容

# 更新字段 ​

## 说明 ​

为该数据表更新字段

## 语法 ​

表达式.UpdateFields(Fields)

表达式: Sheet

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Fields | 是 | Array | 要更新的字段数组 |

## 返回值 ​

Fields

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    await app.Sheets(1).UpdateFields([{ id: 'L', name: '富文本123', type: 'MultiLineText' }]);
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    Application.Sheets(1).UpdateFields([{ id: 'L', name: '富文本123', type: 'MultiLineText' }]);
}
main()
```
 
# 209 API文档 / API / Sheet / OnCreateRecord

本页内容

# 监听新建行记录的事件 ​

## 说明 ​

为当前数据表添加 CreateRecord 事件,当新增 Record 时触发。注意在脚本编辑器中使用时，脚本运行结束就会退出运行，这时可能回调无法被正常触发

## 语法 ​

表达式.OnCreateRecord(Callback)

表达式: Sheet

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Callback | 是 | func | 指定事件发生时的回调方法 ; const eventContext = await Sheet.OnCreateRecord(()=&gt;{ ... }) |

## 返回值 ​

EventContext

## 事件返回数据 ​

RecordRange

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    let eventContext;
    eventContext = await app.Sheets(1).OnCreateRecord(data => {
        console.log(data);
        // 取消事件监听
        eventContext.Destroy();
    });
    app.Sheets(1).Views(1).Records.Add(1, undefined, 5);
    //这里会执行OnCreateRecord的回调
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    let eventContext;
    eventContext = Application.Sheets(1).OnCreateRecord(data => {
        console.log(data);
        // 取消事件监听
        eventContext.Destroy();
    });
    Application.Sheets(1).Views(1).RecordRange.Add(1, undefined, 5);
    //这里会执行OnCreateRecord的回调
}
main();
```
 
# 210 API文档 / API / Sheet / OnDelete

本页内容

- 说明
- 语法
- 参数
- 返回值
- 事件返回数据
- 事件返回数据示例
- 浏览器环境示例
- 脚本编辑器 示例

# 监听删除数据表的事件 ​

## 说明 ​

为当前数据表添加 Delete 事件,当删除 Sheet 时触发。注意在脚本编辑器中使用时，脚本运行结束就会退出运行，这时可能回调无法被正常触发

## 语法 ​

表达式.OnDelete(Callback)

表达式: Sheet

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Callback | 是 | func | 指定事件发生时的回调方法 ; const eventContext = await Sheet.OnDelete(()=&gt;{ ... }) |

## 返回值 ​

EventContext

## 事件返回数据 ​

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| sheetId | Number | 表的 Id |

## 事件返回数据示例 ​

```javascript
{
    sheetId: 2
}
```

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    await app.Sheets.Add({ Type: 'xlEtDataBaseSheet' });
    let eventContext;
    eventContext = await app.Sheets(1).OnDelete(data => {
        console.log(data);
        // 取消事件监听
        eventContext.Destroy();
    });
    app.Sheets(1).Delete(true);
    //这里会执行OnDelete的回调
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    Application.Sheets.Add({ Type: 'xlEtDataBaseSheet' });
    let eventContext;
    eventContext = Application.Sheets(1).OnDelete(data => {
        console.log(data);
        // 取消事件监听
        eventContext.Destroy();
    });
    Application.Sheets(1).Delete(true);
    //这里会执行OnDelete的回调
}
main();
```
 
# 211 API文档 / API / Sheet / OnRename

本页内容

- 说明
- 语法
- 参数
- 返回值
- 事件返回数据
- 浏览器环境示例
- 脚本编辑器 示例

# 监听重命名数据表的事件 ​

## 说明 ​

为当前数据表添加 Rename 事件,当被修改 Name 时触发。注意在脚本编辑器中使用时，脚本运行结束就会退出运行，这时可能回调无法被正常触发

## 语法 ​

表达式.OnRename(Callback)

表达式: Sheet

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Callback | 是 | func | 指定事件发生时的回调方法 ; const eventContext = await Sheet.OnRename(()=&gt;{ ... }) |

## 返回值 ​

EventContext

## 事件返回数据 ​

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| Sheet | Sheet | 表 |
| originValue | String | 原表名 |
| value | String | 现表名 |

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    let eventContext;
    eventContext = await app.Sheets(1).OnRename(data => {
        console.log(data);
        // 取消事件监听
        eventContext.Destroy();
    });
    const sheetName = app.Sheets(1).Name;
    //这里会执行OnRename的回调
    sheetName = 'newName';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    let eventContext;
    eventContext = Application.Sheets(1).OnRename(data => {
        console.log(data);
        // 取消事件监听
        eventContext.Destroy();
    });
    Application.Sheets(1).Name = 'newName';
    //这里会执行OnRename的回调
}
main();
```
 
# 212 API文档 / API / Sheets / Sheets对象

本页内容

# Sheets 数据表集合操作 ​

## 说明 ​

- 注意这里可以获取到三种 sheet 类型
- sheet: 泛指数据表（dbSheet）、仪表盘（dashboardSheet）、说明页面（fpSheet）

## 方法 ​

| 方法 | 说明 |
| --- | --- |
| [Add](/documents/app-integration-dev/guide/dbsheet/Api/Sheets_Add.html) | 新建数据表 |
| [Delete](/documents/app-integration-dev/guide/dbsheet/Api/Sheets_Delete.html) | 删除数据表 |
| [Move](/documents/app-integration-dev/guide/dbsheet/Api/Sheets_Move.html) | 移动数据表 |
| [Item](/documents/app-integration-dev/guide/dbsheet/Api/Sheets_Item.html) | 通过索引位置或者名称获取数据表 |
| [ItemById](/documents/app-integration-dev/guide/dbsheet/Api/Sheets_ItemById.html) | 通过Id 获取 数据表 |
| [GetActiveSheetIndex](/documents/app-integration-dev/guide/dbsheet/Api/Sheets_GetActiveSheetIndex.html) | 获取当前激活数据表的索引位置 |
| [GetActiveSheetId](/documents/app-integration-dev/guide/dbsheet/Api/Sheets_GetActiveSheetId.html) | 获取当前激活数据表的表 Id |

## 事件 ​

| 事件 | 说明 |
| --- | --- |
| [OnCreateSheet](/documents/app-integration-dev/guide/dbsheet/Api/Sheets_OnCreateSheet.html) | 新增数据表时触发 |

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    await app.Sheets.Add({ Type: 'xlEtFlexPaperSheet' });
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  Application.Sheets.Add({ Type: 'xlEtFlexPaperSheet' });
}
main()
```
 
# 213 API文档 / API / Sheets / Add

本页内容

- 说明
- 语法
- 参数
- 参数Config属性详解
- 返回值
- 浏览器环境示例
- 脚本编辑器示例

# 新建数据表 ​

JSSDK: v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

新建数据表到指定位置，Before 和 After 只需要提供一个，另一个填 null 即可

## 语法 ​

表达式.Add(Before, After,Type,Config)

表达式：Sheets

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Before | 否 | number/string | 插入到Before(索引从1开始/数据表名)对应sheet之前 |
| After | 否 | number/string | 插入到After(索引从1开始/数据表名)对应sheet之后 |
| Type | 是 | string | "xlEtFlexPaperSheet"(说明页面)(暂不支持)、"xlEtDataBaseSheet"（数据表）、"xlDbDashBoardSheet"（仪表盘） |
| Config | 否 | object | 数据表专属配置，结构：Config:{ fields : Field[] , name ?: string , views ?: View[] }； |

## 参数Config属性详解 ​

| 属性名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| fields | 否 | Field[] | 字段数组，Field类型结构：{fieldType: FieldType,args: { fieldName: string, fieldWidth: number, listItems?: { value: string, color: number}[], numberFormat?: string, maxRating?: number } } |
| name | 否 | string | 数据表名,默认为‘data1’ |
| views | 否 | View[] | 视图配置数组，View结构：{name: string,type: ViewType}，ViewType的取值为：'Grid'（网格视图）、'Kanban'（看板视图）、'Gallery'（相册视图）、'Form'（表单视图）、'Gantt'（甘特视图）、'Query'（查询视图）或'Calendar'（日历视图）；默认创建'Grid'。暂只支持'Grid'和'Form'。 |

## 返回值 ​

Sheet

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  await app.Sheets.Add(
    null,1,'xlEtDataBaseSheet',
        {
            fields:
                [
                    {fieldType:'SingleLineText',args:{fieldName:'文本',fieldWidth:15}},
                    {fieldType:'MultiLineText',args:{fieldName:'多行文本',fieldWidth:15}},
                    {fieldType:'Date',args:{fieldName:'日期',numberFormat:'yyyy/mm/dd;@',fieldWidth:15}},
                    {fieldType:'SingleSelect',args:{fieldName:'单选项',fieldWidth:15,
                        listItems:[{value: '选项1', color: 4283466178},{value: '选项2',color: 4281378020}]}},
                    {fieldType:'Number',args:{fieldName:'数字',fieldWidth:15}},
                    {fieldType:'Rating',args:{fieldName:'等级',maxRating:6,fieldWidth:15}},
                ],
            name:'数据表',
            views:
                [
                    {name:'表格视图',type:'Grid'},
                    {name:'表单视图',type:'Form'}
                ]
        }
    )
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
  Application.Sheets.Add(
    1,null,'xlEtDataBaseSheet',
        {
            fields:
                [
                    {fieldType:'SingleLineText',args:{fieldName:'文本',fieldWidth:15}},
                    {fieldType:'MultiLineText',args:{fieldName:'多行文本',fieldWidth:15}},
                    {fieldType:'Date',args:{fieldName:'日期',numberFormat:'yyyy/mm/dd;@',fieldWidth:15}},
                    {fieldType:'SingleSelect',args:{fieldName:'单选项',fieldWidth:15,
                        listItems:[{value: '选项1', color: 4283466178},{value: '选项2',color: 4281378020}]}},
                    {fieldType:'Number',args:{fieldName:'数字',fieldWidth:15}},
                    {fieldType:'Rating',args:{fieldName:'等级',maxRating:6,fieldWidth:15}},
                ],
            name:'数据表',
            views:
                [
                    {name:'表格视图',type:'Grid'},
                    {name:'表单视图',type:'Form'}
                ]
        }
    )
}
main()
```
 
# 214 API文档 / API / Sheets / Delete

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 删除数据表 ​

## 说明 ​

通过索引位置或数据表名来删除指定表

## 语法 ​

表达式.Delete(Index)

表达式: Sheets

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Index | 是 | number/string | 索引从 1 开始/数据表名 |

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    await app.Sheets.Delete(1);
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  Application.Sheets.Delete(1);
}
main()
```
 
# 215 API文档 / API / Sheets / GetActiveSheetId

本页内容

# 获取当前激活数据表的表 Id ​

## 说明 ​

获取当前激活数据表的表 Id

## 语法 ​

表达式: GetActiveSheetId()

表达式: Sheets

## 参数 ​

无参数

## 返回值 ​

Number

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const sheetId = await app.Sheets.GetActiveSheetId();
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const sheetId = Application.Sheets.GetActiveSheetId();
}
main()
```
 
# 216 API文档 / API / Sheets / GetActiveSheetIndex

本页内容

# 获取当前激活数据表的索引 ​

## 说明 ​

获取当前激活数据表的索引位置

## 语法 ​

表达式.GetActiveSheetIndex()

表达式: Sheets

## 参数 ​

无参数

## 返回值 ​

Number

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const sheetIndex = await app.Sheets.GetActiveSheetIndex();
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const sheetIndex = Application.Sheets.GetActiveSheetIndex();
}
main()
```
 
# 217 API文档 / API / Sheets / Item

本页内容

# 通过索引位置获取 Sheet ​

JSSDK: v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

通过索引位置或者名称获取 Sheet

## 语法 ​

表达式.Item(Index)

表达式：Sheets

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Index | 是 | number/string | 索引从 1 开始/名称 |

## 返回值 ​

Sheet

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const sheet = await app.Sheets.Item(1)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const sheet = Application.Sheets.Item(1);
}
main()
```
 
# 218 API文档 / API / Sheets / ItemById

本页内容

# 通过Id 获取数据表 ​

JSSDK: v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

通过Id 获取 Sheet

## 语法 ​

表达式.ItemById(Id)

表达式：Sheets

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Id | 是 | number/string | 表格Id |

## 返回值 ​

Sheet

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const sheet = await app.Sheets.ItemById(1)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const sheet = Application.Sheets.ItemById(1);
}
main()
```
 
# 219 API文档 / API / Sheets / Move

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 移动数据表 ​

## 说明 ​

移动数据表到指定位置，Before 和 After 只需要提供一个，另一个填 null 即可

## 语法 ​

表达式.Move(From, Before, After)

表达式: Sheets

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| From | 是 | number/string | 待移动 sheet 的名称或索引号，从 1 开始 |
| Before | 是 | number/string | 移动到 Before(索引从 1 开始/数据表名)对应 sheet 之前 |
| After | 是 | number/string | 移动到 After(索引从 1 开始/数据表名)对应 sheet 之后 |

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    await app.Sheets.Move(111111, null, 22222);
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  Application.Sheets.Move(111111, null, 22222);
}
main()
```
 
# 220 API文档 / API / Sheets / OnCreateSheet

本页内容

- 说明
- 语法
- 参数
- 返回值
- 事件返回数据
- 浏览器环境示例
- 脚本编辑器 示例

# 监听增加数据表的事件 ​

## 说明 ​

为当前数据表集合添加 CreateSheet 事件(当前只支持添加数据表事件，添加说明页和仪表盘不会触发该事件，后续版本更新后支持),当新增 sheet 时触发。注意在脚本编辑器中使用时，脚本运行结束就会退出运行，这时可能回调无法被正常触发

## 语法 ​

表达式: OnCreateSheet(Callback)

表达式: Sheets

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Callback | 是 | func | 指定事件发生时的回调方法 ; const eventContext = await Sheets.OnCreateSheet(()=&gt;{ ... }) |

## 返回值 ​

EventContext

## 事件返回数据 ​

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| Sheet | Sheet | 表 |

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    let eventContext;
    eventContext = await app.Sheets.OnCreateSheet(data => {
        console.log(data);
        // 取消事件监听
        eventContext.Destroy();
    });
    await app.Sheets.Add({ Type: 'xlEtDataBaseSheet' });
    //这里会执行OnCreateSheet的回调
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    let eventContext;
    eventContext = Application.Sheets.OnCreateSheet(data => {
        console.log(data);
        // 取消事件监听
        eventContext.Destroy();
    });
    Application.Sheets.Add({ Type: 'xlEtDataBaseSheet' });
    //这里会执行OnCreateSheet的回调
}
main();
```
 
# 221 API文档 / API / CalendarView / CalendarView对象

本页内容

# CalendarView (对象) ​

## 说明 ​

日历视图，是 [View](/documents/app-integration-dev/guide/dbsheet/Api/View.html) 的子类，View 的方法与属性都可用，下面列出它独有的方法与属性。

## 方法 ​

## 属性 ​

| 属性 | 说明 | 读写说明 |
| --- | --- | --- |
| [BeginField](/documents/app-integration-dev/guide/dbsheet/Api/CalendarView_BeginField.html) | 开始日期字段 | 可读写 |
| [EndField](/documents/app-integration-dev/guide/dbsheet/Api/CalendarView_EndField.html) | 结束日期字段 | 可读写 |
| [TimelineColor](/documents/app-integration-dev/guide/dbsheet/Api/CalendarView_TimelineColor.html) | 事件线颜色 | 可读写 |
| [TimelineColorFollowField](/documents/app-integration-dev/guide/dbsheet/Api/CalendarView_TimelineColorFollowField.html) | 事件线颜色跟随字段 | 可读写 |
| [TimelineColorType](/documents/app-integration-dev/guide/dbsheet/Api/CalendarView_TimelineColorType.html) | 事件线颜色类型 | 可读写 |
| [TitleField](/documents/app-integration-dev/guide/dbsheet/Api/CalendarView_TitleField.html) | 日历视图设置标题字段 | 可读写 |

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    console.log(view.BeginField);
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const view = Application.Sheets(1).Views(1);
    console.log(view.BeginField);
}
main();
```
 
# 222 API文档 / API / CalendarView / BeginField

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 开始日期字段 ​

CalendarView.BeginField(属性)

## 说明 ​

可读写

开始日期字段，值为字段 ID.

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    console.log(view.BeginField);
    view.BeginField = 's';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const view = Application.Sheets(1).Views(1);
    console.log(view.BeginField);
    view.BeginField = 's';
}
main();
```
 
# 223 API文档 / API / CalendarView / EndField

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 结束日期字段 ​

CalendarView.EndField(属性)

## 说明 ​

可读写

结束日期字段，值为字段 ID.

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    console.log(view.EndField);
    view.EndField = 's';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const view = Application.Sheets(1).Views(1);
    console.log(view.EndField);
    view.EndField = 's';
}
main();
```
 
# 224 API文档 / API / CalendarView / TimelineColor

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 时间线颜色 ​

CalendarView.TimelineColor(属性)

## 说明 ​

可读写

事件线颜色，值为Hex格式的RGB颜色值，如：#FF0000。

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    console.log(view.TimelineColor);
    view.TimelineColor = '#97E4E4';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const view = Application.Sheets(1).Views(1);
    console.log(view.TimelineColor);
    view.TimelineColor = '#97E4E4';
}
main();
```
 
# 225 API文档 / API / CalendarView / TimelineColorFollowField

本页内容

# 事件线颜色跟随字段 ​

CalendarView.TimelineColorFollowField(属性)

## 说明 ​

可读写

事件线颜色跟随字段，值为字段Id。

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    console.log(view.TimelineColorFollowField);
    view.TimelineColorFollowField = 'E';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const view = Application.Sheets(1).Views(1);
    console.log(view.TimelineColorFollowField);
    view.TimelineColorFollowField = 'E';
}
main();
```
 
# 226 API文档 / API / CalendarView / TimelineColorType

本页内容

# 事件线颜色类型 ​

CalendarView.TimelineColorType(属性)

## 说明 ​

可读写

事件线颜色类型，值为'Custom'(自定义颜色)或'Follow'(跟随单选项字段)

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    console.log(view.TimelineColorType);
    view.TimelineColor = 'Custom';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const view = Application.Sheets(1).Views(1);
    console.log(view.TimelineColorType);
    view.TimelineColor = 'Follow';
}
main();
```
 
# 227 API文档 / API / CalendarView / TitleField

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 标题设置 ​

CalendarView.TitleField(属性)

## 说明 ​

可读写

日历视图设置标题字段，值为字段 ID.

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const CalendarView = await app.Sheets(1).Views(1);
    console.log(CalendarView.TitleField);
    CalendarView.TitleField = 's';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const CalendarView = Application.Sheets(1).Views(1);
    console.log(CalendarView.TitleField);
    CalendarView.TitleField = 's';
}
main();
```
 
# 228 API文档 / API / CardViewUI / CardViewUI对象

本页内容

# CardViewUI (对象) ​

## 说明 ​

移动端卡片视图的界面属性，只有移动端的表格视图才存在这个对象

## 属性 ​

| 属性 | 说明 | 读写说明 |
| --- | --- | --- |
| [ViewMode](/documents/app-integration-dev/guide/dbsheet/Api/CardViewUI_ViewMode.html) | 卡片视图的模式 | 可读写 |

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const noticeBar = await app.Window.CardViewUI
 }
```
 
# 229 API文档 / API / CardViewUI / ViewMode

本页内容

# 模式 ​

CardViewUI.ViewMode(属性)

## 说明 ​

可读写 卡片视图的模式，注意：设置值时为以下之一：'Grid','Card'

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function today() {
    await instance.ready();
    const app = instance.Application;
    const ViewUI = await app.Window.CardViewUI
    console.log(await ViewUI.ViewMode)
}
```
 
# 230 API文档 / API / GanttView / GanttView对象

本页内容

# GanttView (对象) ​

## 说明 ​

甘特视图，是 [View](/documents/app-integration-dev/guide/dbsheet/Api/View.html) 的子类，View 的方法与属性都可用，下面列出它独有的方法与属性。

## 方法 ​

## 属性 ​

| 属性 | 说明 | 读写说明 |
| --- | --- | --- |
| [BeginField](/documents/app-integration-dev/guide/dbsheet/Api/GanttView_BeginField.html) | 开始日期字段 | 可读写 |
| [EndField](/documents/app-integration-dev/guide/dbsheet/Api/GanttView_EndField.html) | 结束日期字段 | 可读写 |
| [IsOnlyWorkDay](/documents/app-integration-dev/guide/dbsheet/Api/GanttView_IsOnlyWorkDay.html) | 是否忽略节假日 | 可读写 |
| [TimelineColor](/documents/app-integration-dev/guide/dbsheet/Api/GanttView_TimelineColor.html) | 事件线颜色 | 可读写 |
| [TimelineColorFollowField](/documents/app-integration-dev/guide/dbsheet/Api/GanttView_TimelineColorFollowField.html) | 事件线颜色跟随字段 | 可读写 |
| [TimelineColorType](/documents/app-integration-dev/guide/dbsheet/Api/GanttView_TimelineColorType.html) | 事件线颜色类型 | 可读写 |

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const ganttView = await app.Sheets(1).Views(1);
    console.log(ganttView.IsOnlyWorkDay);
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const ganttView = Application.Sheets(1).Views(1);
    console.log(ganttView.IsOnlyWorkDay);
}
main();
```
 
# 231 API文档 / API / GanttView / BeginField

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 开始日期字段 ​

## 说明 ​

可读写

开始日期字段，值为字段 ID.

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const ganttView = await app.Sheets(1).Views(1);
    console.log(ganttView.BeginField);
    ganttView.BeginField = 's';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const ganttView = Application.Sheets(1).Views(1);
    console.log(ganttView.BeginField);
    ganttView.BeginField = 's';
}
main();
```
 
# 232 API文档 / API / GanttView / Calendars

本页内容

# 自定义日历 ​

GanttView.Calendars

## 说明 ​

可读写

自定义日历。

## 返回值 ​

javascript

```javascript
{
    isDefault?: boolean,    // 是否为默认日历
    name?: string,  // 日历名称(暂不支持指定)
    locality?: number, // 地区编码
    weekend?: number[],  // 值为0-6其中之一 (6,0代表周六、周日)
    customWorkday?: number[], // 同上
    customHoliday?: number[], // 同上
    ignoreLawHoliday?: boolean,   // 是否忽略法定节假日
}
```

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const ganttView = await app.Sheets(1).Views(1);
    console.log(ganttView.Calendars);
    ganttView.Calendars =
        {
            isDefault: true,
            name: '自定义日历',
            locality: 156,
            weekend: [6, 0],
            customWorkday: [1, 2, 3, 4, 5],
            customHoliday: [],
            ignoreLawHoliday: true,
        }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const ganttView = Application.Sheets(1).Views(1);
    console.log(ganttView.Calendars);
    ganttView.Calendars = 
        {
            isDefault: true,
            name: '自定义日历',
            locality: 156,
            weekend: [6, 0],
            customWorkday: [1, 2, 3, 4, 5],
            customHoliday: [],
            ignoreLawHoliday: true,
        }
}
main();
```
 
# 233 API文档 / API / GanttView / EndField

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 结束日期字段 ​

## 说明 ​

可读写

结束日期字段，值为字段 ID.

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const ganttView = await app.Sheets(1).Views(1);
    console.log(ganttView.EndField);
    ganttView.EndField = 's';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const ganttView = Application.Sheets(1).Views(1);
    console.log(ganttView.EndField);
    ganttView.EndField = 's';
}
main();
```
 
# 234 API文档 / API / GanttView / IsOnlyWorkDay

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 工时统计-忽略节假日 ​

GanttView.IsOnlyWorkDay

## 说明 ​

可读写

是否忽略节假日。

## 返回值 ​

boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const ganttView = await app.Sheets(1).Views(1);
    console.log(ganttView.IsOnlyWorkDay);
    ganttView.IsOnlyWorkDay = false;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const ganttView = Application.Sheets(1).Views(1);
    console.log(ganttView.IsOnlyWorkDay);
    ganttView.IsOnlyWorkDay = true;
}
main();
```
 
# 235 API文档 / API / GanttView / TimelineColor

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 时间线颜色 ​

GanttView.TimelineColor

## 说明 ​

可读写

时间线颜色，值为Hex格式的RGB颜色值，如：#FF0000。

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const ganttView = await app.Sheets(1).Views(1);
    console.log(ganttView.TimelineColor);
    ganttView.TimelineColor = '#97E4E4';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const ganttView = Application.Sheets(1).Views(1);
    console.log(ganttView.TimelineColor);
    ganttView.TimelineColor = '#97E4E4';
}
main();
```
 
# 236 API文档 / API / GanttView / TimelineColorFollowField

本页内容

# 事件线颜色跟随字段 ​

## 说明 ​

可读写

事件线颜色跟随字段，值为字段Id。

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const ganttView = await app.Sheets(1).Views(1);
    console.log(ganttView.TimelineColorFollowField);
    ganttView.TimelineColorFollowField = 'E';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const ganttView = Application.Sheets(1).Views(1);
    console.log(ganttView.TimelineColorFollowField);
    ganttView.TimelineColorFollowField = 'E';
}
main();
```
 
# 237 API文档 / API / GanttView / TimelineColorType

本页内容

# 事件线颜色类型 ​

## 说明 ​

可读写

事件线颜色类型，值为'Custom'(自定义颜色)或'Follow'(跟随单选项字段)

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const ganttView = await app.Sheets(1).Views(1);
    console.log(ganttView.TimelineColorType);
    ganttView.TimelineColor = 'Custom';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const ganttView = Application.Sheets(1).Views(1);
    console.log(ganttView.TimelineColorType);
    ganttView.TimelineColor = 'Follow';
}
main();
```
 
# 238 API文档 / API / GanttViewUI / GanttViewUI对象

本页内容

# GanttViewUI (对象) ​

## 说明 ​

甘特图界面对象

## 方法 ​

| 方法 | 说明 |
| --- | --- |
| [Today](/documents/app-integration-dev/guide/dbsheet/Api/GanttViewUI_Today.html) | 跳转到今天 |
| [NextPage](/documents/app-integration-dev/guide/dbsheet/Api/GanttViewUI_NextPage.html) | 跳转到下一页 |
| [PrevPage](/documents/app-integration-dev/guide/dbsheet/Api/GanttViewUI_PrevPage.html) | 跳转到前一页 |

## 属性 ​

| 属性 | 说明 | 读写说明 |
| --- | --- | --- |
| [GanttGridFold](/documents/app-integration-dev/guide/dbsheet/Api/GanttViewUI_GanttGridFold.html) | 甘特图字段显示是否折叠 | 可读写 |
| [ViewMode](/documents/app-integration-dev/guide/dbsheet/Api/GanttViewUI_ViewMode.html) | 甘特视图的模式 | 可读写 |

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const noticeBar = await app.Window.GanttViewUI
 }
```
 
# 239 API文档 / API / GanttViewUI / GanttGridFold

本页内容

- 说明
- 返回值
- 浏览器环境示例

# 设置折叠 ​

GanttViewUI.GanttGridFold

## 说明 ​

可读写 设置甘特图字段显示是否折叠

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function today() {
    await instance.ready();
    const app = instance.Application;
    const GanttViewUI = await app.Window.GanttViewUI
    // 设置折叠
    GanttViewUI.GanttGridFold = true
}
```
 
# 240 API文档 / API / GanttViewUI / ViewMode

本页内容

# 模式 ​

GanttViewUI.ViewMode

## 说明 ​

可读写 甘特视图的模式，注意：设置值时为以下之一：'Week','Month','Quarter','Year' .

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function today() {
    await instance.ready();
    const app = instance.Application;
    const GanttViewUI = await app.Window.GanttViewUI
    GanttViewUI.ViewMode = "week"
}
```
 
# 241 API文档 / API / GanttViewUI / NextPage

本页内容

- 说明
- 返回值
- 浏览器环境示例

# 时间线操作-定位下一页 ​

GanttViewUI.NextPage(方法)

## 说明 ​

跳转到下一页

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function today() {
    await instance.ready();
    const app = instance.Application;
    const GanttViewUI = await app.Window.GanttViewUI
    GanttViewUI.NextPage()
}
```
 
# 242 API文档 / API / GanttViewUI / PrevPage

本页内容

- 说明
- 返回值
- 浏览器环境示例

# 时间线操作-定位上一页 ​

GanttViewUI.PrevPage

## 说明 ​

跳转到上一页

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function today() {
    await instance.ready();
    const app = instance.Application;
    const GanttViewUI = await app.Window.GanttViewUI
    GanttViewUI.PrevPage()
}
```
 
# 243 API文档 / API / GanttViewUI / Today

本页内容

- 说明
- 返回值
- 浏览器环境示例

# 时间线操作-定位至今天 ​

GanttViewUI.Today(方法)

## 说明 ​

跳转到今天

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function today() {
    await instance.ready();
    const app = instance.Application;
    const GanttViewUI = await app.Window.GanttViewUI
    GanttViewUI.Today()
}
```
 
# 244 API文档 / API / GridView / GridView对象

本页内容

# GridView (对象) ​

## 说明 ​

表格视图，是 [View](/documents/app-integration-dev/guide/dbsheet/Api/View.html) 的子类，View 的方法与属性都可用，下面列出它独有的方法与属性。

## 方法 ​

## 属性 ​

| 属性 | 说明 | 读写说明 |
| --- | --- | --- |
| [RowHeight](/documents/app-integration-dev/guide/dbsheet/Api/GridView_RowHeight.html) | 表格视图的行高 | 可读写 |
| [FrozenCols](/documents/app-integration-dev/guide/dbsheet/Api/GridView_FrozenCols.html) | 表格视图的冻结列数 | 可读写 |

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const gridView = await app.Sheets(1).Views(1);
    console.log(gridView.RowHeight);
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const gridView = Application.Sheets(1).Views(1);
    console.log(gridView.RowHeight);
}
main();
```
 
# 245 API文档 / API / GridView / FrozenCols

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 冻结列数 ​

GridView.FrozenCols

## 说明 ​

可读写

表格视图的冻结列数

## 返回值 ​

Number

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const gridView = await app.Sheets(1).Views(1);
    console.log(gridView.FrozenCols);
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const gridView = Application.Sheets(1).Views(1);
    console.log(gridView.FrozenCols);
}
main();
```
 
# 246 API文档 / API / GridView / RowHeight

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 行高 ​

GridView.RowHeight

## 说明 ​

可读写

表格视图的行高，注意：设置值时为以下之一：'Short','Medium','Tall','ExtraTall' .

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const gridView = await app.Sheets(1).Views(1);
    console.log(gridView.RowHeight);
    gridView.RowHeight = 'Tall';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const gridView = Application.Sheets(1).Views(1);
    console.log(gridView.RowHeight);
    gridView.RowHeight = 'Tall';
}
main();
```
 
# 247 API文档 / API / QueryView / QueryView对象

本页内容

# QueryView (对象) ​

## 说明 ​

表格视图，是 [View](/documents/app-integration-dev/guide/dbsheet/Api/View.html) 的子类，View 的方法与属性都可用，下面列出它独有的方法与属性。

## 方法 ​

## 属性 ​

| 属性 | 说明 | 读写说明 |
| --- | --- | --- |
| [QueryFields](/documents/app-integration-dev/guide/dbsheet/Api/QueryView_QueryFields.html) | 查询视图的查询条件配置数组 | 可读写 |
| [BackgroundImage](/documents/app-integration-dev/guide/dbsheet/Api/QueryView_BackgroundImage.html) | 查询视图的背景图 | 可读写 |

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    console.log(view.QueryFields);
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const view = Application.Sheets(1).Views(1);
    console.log(view.QueryFields);
}
main();
```
 
# 248 API文档 / API / QueryView / BackgroundImage

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置背景图 ​

QueryView.BackgroundImage

## 说明 ​

可读写

查询视图的背景图，注意：可以设置为 url/base64

## 返回值 ​

Attachment

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    view.BackgroundImage = "https://kdocs-om.wpscdn.cn/om/image.png"
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const view = Application.Sheets(1).Views(1);
    view.BackgroundImage = "https://kdocs-om.wpscdn.cn/om/image.png"
}
main();
```
 
# 249 API文档 / API / QueryView / QueryFields

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置查询条件 ​

QueryView.QueryFields

## 说明 ​

可读写

查询视图的查询条件配置数组，可以将数组设置到QueryFields属性，查询条件的数据结构如下

javascript

```javascript
{
conditionCanBlank: false, // 是否必填
customPrompt: "", // 自定义提示语
enableScanCodeToInput: false,  // 是否允许扫码输入
fieldId: "s",  // 字段ID
needSecondCheck: false,  // 电话字段时是否校验号码
op: "Equals" // 匹配方式，参看下面说明
}
```

根据字段类型支持不同的匹配方式 文本/邮箱/URL/地址/级联：Intersected，Contains，Equals 日期： Intersected，GreaterEquAndLessEqu，Equals 时间： Equals 数字/货币/百分比/最后修改时间/等级/进度/创建时间： GreaterEquAndLessEqu， Equals 身份证/电话/自动编号：Intersected，Equals 复选框/单选项/多选项/联系人/创建人/最后修改人/双向关联/单向关联/父记录：Intersected

## 返回值 ​

Array

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    console.log(view.QueryFields);
    // 添加查询条件
    view.QueryFields = [{
        conditionCanBlank: false, // 是否必填
        customPrompt: "", // 自定义提示语
        enableScanCodeToInput: false,  // 是否允许扫码输入
        fieldId: "s",  // 字段ID
        needSecondCheck: false,  // 电话字段时是否校验号码
        op: "Equals" // 匹配方式，参看下面说明
    }]
    // 使用手机验证码
    view.QueryFields = [{
        conditionCanBlank: false,
        customPrompt: "",
        enableScanCodeToInput: false,
        fieldId: "s",
        needSecondCheck: true
        op: "Equals"
    }]
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const view = Application.Sheets(1).Views(1);
    console.log(view.QueryFields);
}
main();
```
 
# 250 API文档 / API / View / View对象

本页内容

# View (对象) ​

## 说明 ​

视图

## 方法 ​

| 方法 | 说明 |
| --- | --- |
| [Activate](/documents/app-integration-dev/guide/dbsheet/Api/View_Activate.html) | 激活该视图 |
| [Copy](/documents/app-integration-dev/guide/dbsheet/Api/View_Copy.html) | 复制视图 |
| [Delete](/documents/app-integration-dev/guide/dbsheet/Api/View_Delete.html) | 删除视图 |

## 属性 ​

| 属性 | 说明 | 读写说明 |
| --- | --- | --- |
| [Description](/documents/app-integration-dev/guide/dbsheet/Api/View_Description.html) | 视图说明 |  |
| [IsFavView](/documents/app-integration-dev/guide/dbsheet/Api/View_IsFavView.html) | 快速访问视图 |  |
| [Fields](/documents/app-integration-dev/guide/dbsheet/Api/View_Fields.html) | 获取当前视图下的字段列表 | 可读 |
| [Filters](/documents/app-integration-dev/guide/dbsheet/Api/View_Filters.html) | 获取当前视图下的筛选器 | 可读 |
| [Groups](/documents/app-integration-dev/guide/dbsheet/Api/View_Groups.html) | 获取指定视图下面的分组列表 | 可读 |
| [Id](/documents/app-integration-dev/guide/dbsheet/Api/View_Id.html) | 视图Id |  |
| [Name](/documents/app-integration-dev/guide/dbsheet/Api/View_Name.html) | 视图名称 |  |
| [IsPersonal](/documents/app-integration-dev/guide/dbsheet/Api/View_IsPersonal.html) | 是否公共视图 |  |
| [RecordRange](/documents/app-integration-dev/guide/dbsheet/Api/View_RecordRange.html) | 视图的RecordRange |  |
| [Records](/documents/app-integration-dev/guide/dbsheet/Api/View_Records.html) | 视图的Records |  |
| [Selection](/documents/app-integration-dev/guide/dbsheet/Api/View_Selection.html) | 该视图的选中的RecordRange |  |
| [Sorts](/documents/app-integration-dev/guide/dbsheet/Api/View_Sorts.html) | 获取或设置当前视图下的排序条件列表 | 可读写 |
| [ViewShare](/documents/app-integration-dev/guide/dbsheet/Api/View_ViewShare.html) | 该视图的视图分享 |  |

## 事件 ​

| 事件 | 说明 |
| --- | --- |
| [OnDelete](/documents/app-integration-dev/guide/dbsheet/Api/View_OnDelete.html) |  |
| [OnRename](/documents/app-integration-dev/guide/dbsheet/Api/View_OnRename.html) |  |

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const view = Application.Sheets(1).Views(1);
}
main()
```
 
# 251 API文档 / API / View / Description

本页内容

# View.Description(属性) ​

## 说明 ​

视图说明

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
// 获取视图说明
async function getDescription() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    const description = await view.Description;
}

// 设置视图说明
async function setDescription() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    view.Description = '新视图说明';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const view = Application.Sheets(1).Views(1);
  const description = view.Description;
  view.Description = '新视图说明';
}
main()
```
 
# 252 API文档 / API / View / Fields

本页内容

# View.Fields(属性) ​

## 说明 ​

可读 获取当前视图下的字段列表

## 返回值 ​

Fields

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const fields = await app.Sheets(1).Views(1).Fields;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const fields = Application.Sheets(1).Views(1).Fields;
}
main()
```
 
# 253 API文档 / API / View / Filters

本页内容

# View.Filters(属性) ​

## 说明 ​

可读 获取当前视图下的筛选器

## 返回值 ​

Filters

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const filters = await app.Sheets(1).Views(1).Filters;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const filters = Application.Sheets(1).Views(1).Filters;
}
main()
```
 
# 254 API文档 / API / View / Groups

本页内容

# View.Groups(属性) ​

## 说明 ​

可读 获取指定视图下面的分组列表

## 返回值 ​

Groups

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    const groups = await view.Groups;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const view = Application.Sheets(1).Views(1);
  const groups = view.Groups;
}
main()
```
 
# 255 API文档 / API / View / Id

本页内容

# View.Id(属性) ​

## 说明 ​

视图Id

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    const viewId = await view.Id;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const viewId = Application.Sheets(1).Views(1).Id;
}
main()
```
 
# 256 API文档 / API / View / IsFavView

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置快速访问视图 ​

View.IsFavView

## 说明 ​

可读写 设置视图是否为快速访问视图

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function getFavView() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    const isFavView = await view.IsFavView; // 若为快速访问视图返回true，若不为快速访问视图返回false
}

async function setFavView() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    view.IsFavView = false; // 取消设置为快速访问视图
    view.IsFavView = true; // 设置为快速访问视图
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function getFavView() {
    const view = Application.Sheets(1).Views(1);
    const isFavView = view.IsFavView; // 若为快速访问视图返回true，若不为快速访问视图返回false
}
function setFavView() {
    const view = Application.Sheets(1).Views(1);
    view.IsFavView = false; // 取消设置为快速访问视图
    view.IsFavView = true; // 设置为快速访问视图
}
main();
```
 
# 257 API文档 / API / View / IsPersonal

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置个人/公共视图 ​

View.IsPersonal

## 说明 ​

可读写 设置视图为个人视图或公共视图

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function getPublicView() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    const isPersonal = await view.IsPersonal; // 若为个人视图返回true，若不为公共视图返回false
    return !isPersonal;
}

async function setPublicView() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    view.IsPersonal = false; // 设置为公共视图
    view.IsPersonal = true; //  设置为个人视图
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const view = Application.Sheets(1).Views(1);
  const isPublicView = !view.IsPersonal;
  view.IsPersonal = false; // 设置为公共视图
  view.IsPersonal = true; // 设置为个人视图
}
main()
```
 
# 258 API文档 / API / View / Name

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 重命名 ​

## 说明 ​

重命名、 获取视图名称

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
// 获取视图名称
async function getName() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    const name = await view.Name;
}
// 设置视图名称
async function setName() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    view.Name = '新视图名称';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const view = Application.Sheets(1).Views(1);
  const name = view.Name;
  view.Name = '新视图名称';
}
main()
```
 
# 259 API文档 / API / View / RecordRange

本页内容

# View.RecordRange(属性) ​

## 说明 ​

视图的RecordRange

## 返回值 ​

RecordRange

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    const recordRange = await view.RecordRange;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const recordRange = Application.Sheets(1).Views(1).RecordRange;
}
main()
```
 
# 260 API文档 / API / View / Records

本页内容

# View.Records(属性) ​

## 说明 ​

视图的Records

## 返回值 ​

Records

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    const records = await view.Records;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const records = Application.Sheets(1).Views(1).Records;
}
main()
```
 
# 261 API文档 / API / View / Selection

本页内容

# View.Selection(属性) ​

## 说明 ​

该视图的选中的RecordRange

## 返回值 ​

RecordRange

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    const recordRange = await view.Selection;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const recordRange = Application.Sheets(1).Views(1).Selection;
}
main()
```
 
# 262 API文档 / API / View / Sorts

本页内容

# View.Sorts(属性) ​

## 说明 ​

可读写 获取或设置当前视图下的排序条件列表

## 返回值 ​

Sorts

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const sorts = await app.Sheets(1).Views(1).Sorts;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const sorts = Application.Sheets(1).Views(1).Sorts;
}
main()
```
 
# 263 API文档 / API / View / Type

本页内容

# View.Type(属性) ​

## 说明 ​

只读属性，视图类型，视图类型包括以下几种 'Grid', // 网格视图 'Kanban', // 看板视图 'Gallery', // 画册视图 'Form', // 表单视图 'Gantt', // 甘特视图 'Query', // 查询视图 'Calendar', // 日历视图

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
// 获取视图名称
async function main() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    const type = await view.Type;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const view = Application.Sheets(1).Views(1);
  const type = view.Type;
}
main()
```
 
# 264 API文档 / API / View / ViewShare

本页内容

# View.ViewShare(属性) ​

## 说明 ​

该视图的视图分享

## 返回值 ​

ViewShare

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    const viewShare = await view.ViewShare;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const viewShare = Application.Sheets(1).Views(1).ViewShare;
}
main()
```
 
# 265 API文档 / API / View / Activate

本页内容

# View.Activate(方法) ​

## 说明 ​

激活该视图

## 语法 ​

表达式.Activate()

表达式: View

## 参数 ​

## 返回值 ​

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    await app.Sheets(1).Views(1).Activate();
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    Application.Sheets(1).Views(1).Activate();
}
main();
```
 
# 266 API文档 / API / View / Copy

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 复制视图 ​

## 说明 ​

复制视图

## 语法 ​

表达式.Copy()

表达式:View

## 参数 ​

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    const result = await view.Copy();
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const view = Application.Sheets(1).Views(1);
  const result = view.Copy();
}
main()
```
 
# 267 API文档 / API / View / Delete

本页内容

- 说明
- 语法
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 删除视图 ​

## 说明 ​

删除指定视图

## 语法 ​

表达式.Delete()

表达式:View

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views(1);
    await view.Delete();
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const view = Application.Sheets(1).Views(1);
  view.Delete();
}
main()
```
 
# 268 API文档 / API / View / OnDelete

本页内容

- 说明
- 语法
- 参数
- 返回值
- 事件返回数据
- 事件返回数据示例
- 浏览器环境示例
- 脚本编辑器 示例

# 监听删除视图的事件 ​

View.OnDelete(方法)

## 说明 ​

为当前视图添加 Delete 事件,当删除 View 时触发。注意在脚本编辑器中使用时，脚本运行结束就会退出运行，这时可能回调无法被正常触发

## 语法 ​

表达式.OnDelete(Callback)

表达式: View

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Callback | 是 | func | 指定事件发生时的回调方法 ; const eventContext = await View.OnDelete(()=&gt;{ ... }) |

## 返回值 ​

EventContext

## 事件返回数据 ​

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| sheetId | Number | 表的 Id |
| viewId | String | 视图的 Id |

## 事件返回数据示例 ​

```javascript
{
    sheetId: 1
    viewId: 'B'
}
```

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    let eventContext;
    eventContext = await app
        .Sheets(1)
        .Views(1)
        .OnDelete(data => {
            console.log(data);
            // 取消事件监听
            eventContext.Destroy();
        });
    app.Sheets(1).Views(1).Delete();
    //这里会执行OnDelete的回调
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    let eventContext;
    eventContext = Application.Sheets(1)
        .Views(1)
        .OnDelete(data => {
            console.log(data);
            // 取消事件监听
            eventContext.Destroy();
        });
    Application.Sheets(1).Views(1).Delete();
    //这里会执行OnDelete的回调
}
main();
```
 
# 269 API文档 / API / View / OnRename

本页内容

- 说明
- 语法
- 参数
- 返回值
- 事件返回数据
- 事件返回数据示例
- 浏览器环境示例
- 脚本编辑器 示例

# 监听重命名视图的事件 ​

View.OnRename(方法)

## 说明 ​

为当前视图添加 Rename 事件,当被修改 Name 时触发。注意在脚本编辑器中使用时，脚本运行结束就会退出运行，这时可能回调无法被正常触发

## 语法 ​

表达式.OnRename(Callback)

表达式: View

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Callback | 是 | func | 指定事件发生时的回调方法 ; const eventContext = await View.OnRename(()=&gt;{ ... }) |

## 返回值 ​

EventContext

## 事件返回数据 ​

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| View | View | 视图 |
| originValue | String | 原始值 |
| value | String | 修改后的值 |

## 事件返回数据示例 ​

View, originValue, value

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    let eventContext;
    eventContext = await app
        .Sheets(1)
        .Views(1)
        .OnRename(data => {
            console.log(data);
            // 取消事件监听
            eventContext.Destroy();
        });
    const sheetName = app.Sheets(1).Views(1).Name;
    sheetName = 'newName';
    //这里会执行OnRename的回调
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    let eventContext;
    eventContext = Application.Sheets(1)
        .Views(1)
        .OnRename(data => {
            console.log(data);
            // 取消事件监听
            eventContext.Destroy();
        });
    Application.Sheets(1).Views(1).Name = 'newName';
    //这里会执行OnRename的回调
}
main();
```
 
# 270 API文档 / API / ViewShare / ViewShare对象

本页内容

# ViewShare (对象) ​

## 说明 ​

单个视图的分享视图相关

## 方法 ​

| 属性 | 说明 |
| --- | --- |
| [SetEnable](/documents/app-integration-dev/guide/dbsheet/Api/ViewShare_SetEnable.html) | 切换当前视图的分享视图开关 |
| [ChangePermission](/documents/app-integration-dev/guide/dbsheet/Api/ViewShare_ChangePermission.html) | 修改分享视图的权限 |

## 属性 ​

| 属性 | 说明 | 读写说明 |
| --- | --- | --- |
| [CanAddRemoveRecords](/documents/app-integration-dev/guide/dbsheet/Api/ViewShare_CanAddRemoveRecords.html) | 返回当前分享视图高级权限中，是否允许添加记录 | 可读写 |
| [VisibleRecType](/documents/app-integration-dev/guide/dbsheet/Api/ViewShare_VisibleRecType.html) | 单个视图的分享视图相关 |  |
| [EditableRecType](/documents/app-integration-dev/guide/dbsheet/Api/ViewShare_EditableRecType.html) | 返回当前分享视图高级权限中，可编辑的记录范围 | 可读 |
| [SharedLinkInfo](/documents/app-integration-dev/guide/dbsheet/Api/ViewShare_SharedLinkInfo.html) | 返回当前分享视图的信息 | 可读 |
| [ShareUrl](/documents/app-integration-dev/guide/dbsheet/Api/ViewShare_ShareUrl.html) | 返回当前分享视图的分享链接 | 可读 |
| [EditableFieldsInfo](/documents/app-integration-dev/guide/dbsheet/Api/ViewShare_EditableFieldsInfo.html) | 返回当前分享视图高级权限中，允许编辑的字段id | 可读写 |

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const viewShare = await app.Sheets(1).Views(1).ViewShare;
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const viewShare = Application.Sheets(1).Views(1).ViewShare;
}
main()
```
 
# 271 API文档 / API / ViewShare / CanAddRemoveRecords

本页内容

# ViewShare.CanAddRecords(属性) ​

## 说明 ​

可读写 返回当前分享视图高级权限中，是否允许添加记录

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const viewShare = await app.Sheets(1).Views(1).ViewShare
    // read
    const canAddRecords = await viewShare.CanAddRemoveRecords;
    // write
    viewShare.CanAddRemoveRecords = true;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const viewShare = Application.Sheets(1).Views(1).ViewShare
    // read
    const canAddRecords = viewShare.CanAddRemoveRecords;
    // write
    viewShare.CanAddRemoveRecords = true;
}
main()
```
 
# 272 API文档 / API / ViewShare / EditableFieldsInfo

本页内容

# ViewShare.EditableFieldIds(属性) ​

## 说明 ​

可读写 返回当前分享视图高级权限中，允许编辑的字段id

## 返回值 ​

Object

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const view = instance.Application.Sheets(1).Views(1);
    const viewShare = view.ViewShare
    
    // read
    const editableFieldsInfo = await viewShare.EditableFieldsInfo;
    
    // write
    // 所有字段可编辑
    viewShare.EditableFieldsInfo = {type: 'All'}

    // 当前视图中索引为1的字段，在分享视图中可编辑
    const fields = view.Fields;
    const fieldId = await fields(1).Id
    viewShare.EditableFieldsInfo = {
        type: 'Custom',
        ids: [fieldId]
    }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const view = Application.Sheets(1).Views(1);
    const viewShare = view.ViewShare
    
    // read
    const editableFieldsInfo = viewShare.EditableFieldsInfo;
    
    // write
    // 所有字段可编辑
    viewShare.EditableFieldsInfo = {type: 'All'}
    
    // 当前视图中索引为1的字段，在分享视图中可编辑
    const fields = view.Fields;
    const fieldId = fields(1).Id
    viewShare.EditableFieldsInfo = {
        type: 'Custom',
        ids: [fieldId]
    }
}
main()
```
 
# 273 API文档 / API / ViewShare / EditableRecType

本页内容

# ViewShare.EditableRecType(属性) ​

## 说明 ​

可读 返回当前分享视图高级权限中，可编辑的记录范围

## 返回值 ​

[DbSharedCriteriaType](/documents/app-integration-dev/guide/dbsheet/Api/Enum_DbSharedCriteriaType.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const viewShare = await app.Sheets(1).Views(1).ViewShare;
    // read
    const visible = await viewShare.EditableRecType;
    // write
    viewShare.EditableRecType = 'Self';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const viewShare = Application.Sheets(1).Views(1).ViewShare;
    // read
    const visible = viewShare.EditableRecType;
    // write
    viewShare.EditableRecType = 'Self';
}
main();
```
 
# 274 API文档 / API / ViewShare / ShareUrl

本页内容

# ViewShare.ShareUrl(属性) ​

## 说明 ​

可读 返回当前分享视图的分享链接

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const viewShareUrl = await app.Sheets(1).Views(1).ViewShare.ShareUrl;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const viewShareUrl = Application.Sheets(1).Views(1).ViewShare.ShareUrl;
}
main()
```
 
# 275 API文档 / API / ViewShare / SharedLinkInfo

本页内容

# ViewShare.SharedLinkInfo(属性) ​

## 说明 ​

可读 返回当前分享视图的信息

## 返回值 ​

Object

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const sharedLinkInfo = await app.Sheets(1).Views(1).ViewShare.SharedLinkInfo;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const sharedLinkInfo = Application.Sheets(1).Views(1).ViewShare.SharedLinkInfo;
}
main()
```
 
# 276 API文档 / API / ViewShare / VisibleRecType

本页内容

# ViewShare.VisibleRecType(属性) ​

## 说明 ​

可读写 返回当前分享视图高级权限中，可查看的记录范围

## 返回值 ​

[DbSharedCriteriaType](/documents/app-integration-dev/guide/dbsheet/Api/Enum_DbSharedCriteriaType.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const viewShare = await app.Sheets(1).Views(1).ViewShare;
    // read
    const visible = await viewShare.VisibleRecType;
    // write
    viewShare.VisibleRecType = 'Self';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const viewShare = Application.Sheets(1).Views(1).ViewShare;
    // read
    const visible = viewShare.VisibleRecType;
    // write
    viewShare.VisibleRecType = 'Self';
}
main();
```
 
# 277 API文档 / API / ViewShare / ChangePermission

本页内容

# ViewShare.ChangePermission(方法) ​

## 说明 ​

修改分享视图的权限

## 语法 ​

表达式.ChangePermission(RangeType, PermissionType)

表达式: ViewShare

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| RangeType | 是 | [SharedLinkToType](/documents/app-integration-dev/guide/dbsheet/Api/Enum_SharedLinkToType.html) | 分享的协作者范围 |
| PermissionType | 是 | [SharedLinkPermissionType](/documents/app-integration-dev/guide/dbsheet/Api/Enum_SharedLinkPermissionType.html) | 分享的协作者权限 |

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const viewShare = await app.Sheets(1).Views(1).ViewShare;
    const res = viewShare.ChangePermission('assigned', 'edit')
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const viewShare = Application.Sheets(1).Views(1).ViewShare;
    const res = viewShare.ChangePermission('assigned', 'edit')
}
main()
```
 
# 278 API文档 / API / ViewShare / SetEnable

本页内容

# ViewShare.SetEnable(方法) ​

## 说明 ​

切换当前视图的分享视图开关

## 语法 ​

表达式.SetEnable(Enable)

表达式: ViewShare

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Enable | 是 | boolean | 切换分享视图的开关，true就是开，false就是关 |

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const viewShare = await app.Sheets(1).Views(1).ViewShare;
    const res = viewShare.SetEnable(true)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const viewShare = Application.Sheets(1).Views(1).ViewShare;
    const res = viewShare.SetEnable(true)
}
main()
```
 
# 279 API文档 / API / Views / Views对象

本页内容

# Views (对象) ​

## 说明 ​

视图集合类

## 方法 ​

| 属性 | 说明 |
| --- | --- |
| [Add](/documents/app-integration-dev/guide/dbsheet/Api/Views_Add.html) | 添加视图 |
| [Delete](/documents/app-integration-dev/guide/dbsheet/Api/Views_Delete.html) | 删除视图 |
| [Item](/documents/app-integration-dev/guide/dbsheet/Api/Views_Item.html) | 根据 索引或名称 获取视图 |
| [ItemById](/documents/app-integration-dev/guide/dbsheet/Api/Views_ItemById.html) | 根据 视图Id 获取视图 |

## 属性 ​

| 属性 | 说明 | 读写说明 |
| --- | --- | --- |
| [Count](/documents/app-integration-dev/guide/dbsheet/Api/Views_Count.html) | 获取视图集合的个数 |  |
| [ActiveView](/documents/app-integration-dev/guide/dbsheet/Api/Views_ActiveView.html) | 获取当前激活的视图 |  |

## 事件 ​

| 属性 | 说明 |
| --- | --- |
| [OnCreate](/documents/app-integration-dev/guide/dbsheet/Api/Views_OnCreate.html) | 添加 View 时触发 |

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const views = await app.Sheets(1).Views;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const views = Application.Sheets(1).Views;
}
main()
```
 
# 280 API文档 / API / Views / ActiveView

本页内容

# Views.ActiveView(属性) ​

## 说明 ​

获取当前激活的视图

## 返回值 ​

View

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const view = await app.Sheets(1).Views.ActiveView;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const view = Application.Sheets(1).Views.ActiveView;
}
main()
```
 
# 281 API文档 / API / Views / Count

本页内容

# Views.Count(属性) ​

## 说明 ​

获取视图集合的个数

## 返回值 ​

Number

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const count = await app.Sheets(1).Views.Count;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const count = Application.Sheets(1).Views.Count;
}
main()
```
 
# 282 API文档 / API / Views / Add

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 新建视图 ​

## 说明 ​

添加视图

## 语法 ​

表达式.Add(Type,Name)

表达式:Views

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Type | 是 | 'Grid'或'Kanban'或'Gallery'或'Form'或’Query‘或‘Gantt’ | 视图类别。 |
| Name | 是 | string | 视图名称 |

## 返回值 ​

[View](/documents/app-integration-dev/guide/dbsheet/Api/View.html), [GridView](/documents/app-integration-dev/guide/dbsheet/Api/GridView.html), [GanttView](/documents/app-integration-dev/guide/dbsheet/Api/GanttView.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const views = await app.Sheets(1).Views;

    const gridView = await views.Add('Grid', '表格视图');
    console.log(await gridView.RowHeight);
    await views.Add('Kanban', '看板视图');
    await views.Add('Gallery', '画册视图');
    await views.Add('Form', '表单视图');
    await views.Add('Query', '查询视图');
    const ganttView = await views.Add('Gantt', '甘特视图');
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const views = Application.Sheets(1).Views;
  const gridView = views.Add('Grid', '表格视图');
  console.log(gridView.RowHeight);
  views.Add('Kanban', '看板视图');
  views.Add('Gallery', '画册视图');
  views.Add('Form', '表单视图');
  views.Add('Query', '查询视图');
  const ganttView = views.Add('Gantt', '甘特视图');
}
main()
```
 
# 283 API文档 / API / Views / Delete

本页内容

# Views.Delete(方法) ​

## 说明 ​

根据索引/Id 删除视图

## 语法 ​

表达式.Delete(Index)

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Index | 是 | number/string | 索引从 1 开始/视图ID |

## 返回值 ​

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const views = await app.Sheets(1).Views;
    await views.Delete(1);
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const views = Application.Sheets(1).Views;
  views.Delete(1);
}
main()
```
 
# 284 API文档 / API / Views / Item

本页内容

# Views.Item(方法) ​

## 说明 ​

根据 索引或名称 获取视图 [View](/documents/app-integration-dev/guide/dbsheet/Api/View.html), 注意当获取到的视图类型不同时，返回的View对象也会不同，参看返回值

## 语法 ​

表达式.Item(Index) 表达式: Views

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Index | 是 | number/string | 索引从 1 开始/ 视图名称 |

## 返回值 ​

[View](/documents/app-integration-dev/guide/dbsheet/Api/View.html) / [GridView](/documents/app-integration-dev/guide/dbsheet/Api/GridView.html)/ [GanttView](/documents/app-integration-dev/guide/dbsheet/Api/GanttView.html)/ [QueryView](/documents/app-integration-dev/guide/dbsheet/Api/QueryView.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const views = await app.Sheets(1).Views;
    const view = await views.Item(1);
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const view = Application.Sheets(1).Views.Item(1);
}
main()
```
 
# 285 API文档 / API / Views / ItemById

本页内容

# Views.ItemById(方法) ​

## 说明 ​

根据 视图Id 获取视图 [View](/documents/app-integration-dev/guide/dbsheet/Api/View.html), 注意当获取到的视图类型不同时，返回的View对象也会不同，参看返回值

## 语法 ​

表达式.ItemById(Id) 表达式: Views

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Id | 是 | string | 视图Id |

## 返回值 ​

[View](/documents/app-integration-dev/guide/dbsheet/Api/View.html) / [GridView](/documents/app-integration-dev/guide/dbsheet/Api/GridView.html)/ [GanttView](/documents/app-integration-dev/guide/dbsheet/Api/GanttView.html)/ [QueryView](/documents/app-integration-dev/guide/dbsheet/Api/QueryView.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const views = await app.Sheets(1).Views;
    const view = await views.ItemById("A");
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const view = Application.Sheets(1).Views.ItemById("A");
}
main()
```
 
# 286 API文档 / API / Views / OnCreate

本页内容

- 说明
- 语法
- 参数
- 返回值
- 事件返回数据
- 浏览器环境示例
- 脚本编辑器 示例

# 监听增加视图的事件 ​

Views.OnCreate(方法)

## 说明 ​

为 Views 添加 Create 事件,当添加 View 时触发。注意在脚本编辑器中使用时，脚本运行结束就会退出运行，这时可能回调无法被正常触发

## 语法 ​

表达式.OnCreate(Callback)

表达式: Views

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Callback | 是 | func | 指定事件发生时的回调方法 ; const eventContext = await Views.OnCreate(()=&gt;{ ... }) |

## 返回值 ​

EventContext

## 事件返回数据 ​

View

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    let eventContext;
    eventContext = await app.Sheets(1).Views.OnCreate(data => {
        console.log(data);
        // 取消事件监听

        eventContext.Destroy();
    });

    await app.Sheets(1).Views.Add('Grid', '表格视图');
    //这里会执行OnCreate的回调
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    let eventContext;
    eventContext = Application.Sheets(1).Views.OnCreate(data => {
        console.log(data);
        // 取消事件监听
        eventContext.Destroy();
    });
    Application.Sheets(1).Views.Add('Grid', '表格视图');
    //这里会执行OnCreate的回调
}
main();
```

# 287 API文档 / API / Record / Record对象

本页内容

# Record (对象) ​

## 说明 ​

Record 对象，表示单条记录相关,Record对象可以返回值的数组或者显示的文本数据，值的数组不同的字段会有不同的对象显示 常见对象如下：

| 字段类型 | 数据显示 |
| --- | --- |
| 文本类型 | String |
| 数值类型 | Number |
| 联系人类型 | {"userId":"10000", "nickname":"name", "avatarUrl":"https://xxxxx/thumbnail/180x180!", "companyId":""} |
| 富文本类型 |  |
| 超链接类型 | {"display":"超链接文本","address":"https://xxx.com"} |
| 多选项类型 | Array |
| 地址类型 |  |
| 关联类型 | [{"id":"B","str":"aaaa"}, {"id":"b","str":"set"}] |
| CheckBox | boolean |
| 等级类型 | Number |
| 级联类型 |  |
| 附件类型 | Attachment |

## 方法 ​

- [Item](/documents/app-integration-dev/guide/dbsheet/Api/Record_Item.html)
- [Delete](/documents/app-integration-dev/guide/dbsheet/Api/Record_Delete.html)
- [Select](/documents/app-integration-dev/guide/dbsheet/Api/Record_Select.html)

## 属性 ​

- [Id](/documents/app-integration-dev/guide/dbsheet/Api/Record_Id.html)
- [Value](/documents/app-integration-dev/guide/dbsheet/Api/Record_Value.html)
- [Text](/documents/app-integration-dev/guide/dbsheet/Api/Record_Text.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const record = await app.Sheets(1).Views(1).Records(10)
    console.log(await record.Value)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const record = Application.Sheets(1).Views(1).Records(10)
    // 返回字段的值
    console.log(record.Value)
 }
main()
```
 
# 288 API文档 / API / Record / Id

本页内容

# Record.Id(属性) ​

## 说明 ​

可读 返回整行记录、单个记录字段的Id

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const record = app.Sheets(1).Views(1).Records(10);
    console.log(await record.Id); 
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const record = Application.Sheets(1).Views(1).Records(10);
    console.log(record.Id); 
}
main()
```
 
# 289 API文档 / API / Record / Text

本页内容

# Record.Text(属性) ​

## 说明 ​

可读 返回整行记录、单个记录字段的值

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const record = app.Sheets(1).Views(1).Records(10)
    console.log(await record.Text) // 获取第10行的记录
    
    const info = app.Sheets(1).Views(1).Records(10, 1)
    console.log(await info.Text) // 获取第10行，第一个字段的值
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const groups = Application.Sheets(1).Views(1).Groups;
}
main()
```
 
# 290 API文档 / API / Record / Value

本页内容

# Record.Value(属性) ​

## 说明 ​

可读写

返回整行记录，抑或返回、设置单个记录字段的值。

## 返回值 ​

ApiResult

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const record = app.Sheets(1).Views(1).Records(10)
    console.log(await record.Value) // 获取第10行的记录
    
    const info = app.Sheets(1).Views(1).Records(10, 1)
    console.log(await info.Value) // 获取第10行，第一个字段的值

    info.Value = '1' // 值设为 1
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const record = Application.Sheets(1).Views(1).Records(10)
    console.log(record.Value) // 获取第10行的记录
    
    const info = Application.Sheets(1).Views(1).Records(10, 1)
    console.log(info.Value) // 获取第10行，第一个字段的值

    info.Value = '1' // 值设为 1
}
main()
```
 
# 291 API文档 / API / Record / Delete

本页内容

# Record.Delete(方法) ​

## 说明 ​

删除当前行

## 语法 ​

表达式.Delete()

表达式：Record

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const record = app.Sheets(1).Views(1).Records(10)
    record.Delete()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const record = Application.Sheets(1).Views(1).Records(10)
  record.Delete()
}
main()
```
 
# 292 API文档 / API / Record / Item

本页内容

# Record.Item(方法) ​

## 说明 ​

获取指定索引位置的字段记录

## 语法 ​

表达式.Item(Index)

表达式：Record

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Index | 是 | number/string | 索引从1开始/字段ID |

## 返回值 ​

Record

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const record = app.Sheets(1).Views(1).Records(10)
    const info = record.Item(1) // 获取第10行号的第一个字段
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const record = Application.Sheets(1).Views(1).Records(10)
  const info = record.Item(1) // 获取第10行号的第一个字段
}
main()
```
 
# 293 API文档 / API / Record / Select

本页内容

- 说明
- 语法
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 选中记录 ​

## 说明 ​

选中某个区域

## 语法 ​

表达式.Select()

表达式：Record

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const record = app.Sheets(1).Views(1).Records(10)
    record.Select() // 选中第10行的记录
    
    const info = app.Sheets(1).Views(1).Records(10, 1)
    await info.Select() // 选中第10行，第一个字段
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const record = Application.Sheets(1).Views(1).Records(10)
    record.Select() // 选中第10行的记录
    
    const info = Application.Sheets(1).Views(1).Records(10, 1)
    info.Select() // 选中第10行，第一个字段
}
main()
```
 
# 294 API文档 / API / RecordRange / RecordRange对象

本页内容

# RecordRange (对象) ​

## 说明 ​

RecordRange对象代表记录Record,如果RecordRange在Sheet上返回，则此RecordRange指向的是Sheet上的数据，如果RecordRange在View上返回，则此RecordRange指向的是视图上的数据。我们为RecordRange提供了多种读取数据的方法。

字段选择器可以使用以下几种来选取记录

1、通过数值指定索引的记录

```javascript
RecordRange(1)
```

2、通过指定记录ID来指定记录

```javascript
RecordRange("a")
```

3、通过符号:可以指定索引范围的记录

```javascript
RecordRange("1:100")
```

4、通过数组可以指定不连续的记录

```javascript
RecordRange(["1","5","10:20"])
```

5、通过参数2为字符串指定字段ID可以限定显示的字段

```javascript
RecordRange(1, "a")
```

6、通过参数2为数组可以指定多个字段ID可以限定显示的字段

```javascript
RecordRange(1, ["a","b","A"])
```

7、通过参数2如果以@字符作为首字符表示字段名称

```javascript
RecordRange(1, ["@名称", "@数量"])
```

8、通过 [Condition()](/documents/app-integration-dev/guide/dbsheet/Api/RecordRange_Condition.html) 方法来筛选符合条件的记录

## 方法 ​

- [Item](/documents/app-integration-dev/guide/dbsheet/Api/RecordRange_Item.html)
- [Add](/documents/app-integration-dev/guide/dbsheet/Api/RecordRange_Add.html)
- [Condition](/documents/app-integration-dev/guide/dbsheet/Api/RecordRange_Condition.html)
- [Delete](/documents/app-integration-dev/guide/dbsheet/Api/RecordRange_Delete.html)
- [Select](/documents/app-integration-dev/guide/dbsheet/Api/RecordRange_Select.html)
- [SetValues](/documents/app-integration-dev/guide/dbsheet/Api/RecordRange_SetValues.html)

## 属性 ​

- [Id](/documents/app-integration-dev/guide/dbsheet/Api/RecordRange_Id.html)
- [Index](/documents/app-integration-dev/guide/dbsheet/Api/RecordRange_Index.html)
- [FieldId](/documents/app-integration-dev/guide/dbsheet/Api/RecordRange_FieldId.html)
- [Count](/documents/app-integration-dev/guide/dbsheet/Api/RecordRange_Count.html)
- [Text](/documents/app-integration-dev/guide/dbsheet/Api/RecordRange_Text.html)
- [Value](/documents/app-integration-dev/guide/dbsheet/Api/RecordRange_Value.html)
- [Font](/documents/app-integration-dev/guide/dbsheet/Api/RecordRange_Font.html)
- [Interior](/documents/app-integration-dev/guide/dbsheet/Api/RecordRange_Interior.html)

## 事件 ​

- [OnUpdate](/documents/app-integration-dev/guide/dbsheet/Api/RecordRange_OnUpdate.html)
- [OnDeleteRecord](/documents/app-integration-dev/guide/dbsheet/Api/RecordRange_OnDeleteRecord.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const values = await app.ActiveView.RecordRange(1, ["@名称", "@日期", "@数量"]).Value
 console.log(values) // 输出 (3) ['aaaa', 45470, 1]
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
 const values = Application.ActiveView.RecordRange(1, ["@名称", "@日期", "@数量"]).Value
 console.log(values) // 输出 (3) ['aaaa', 45470, 1]
}
main()
```
 
# 295 API文档 / API / RecordRange / Count

本页内容

# RecordRange.Count(属性) ​

## 说明 ​

可读写

返回当前RecordRange的记录数量，在View里返回View中可见记录数量，在Sheet返回所有的记录数量，如果指定了 RecordRange的范围则返回RecordRange的记录数量

## 返回值 ​

Number

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const viewCount = await instance.Application.ActiveView.RecordRange.Count
  const sheetCount = await instance.Application.ActiveSheet.RecordRange.Count
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const viewCount = Application.ActiveView.RecordRange.Count
  const sheetCount = Application.ActiveSheet.RecordRange.Count
}
main()
```
 
# 296 API文档 / API / RecordRange / FieldId

本页内容

# RecordRange.FieldId(属性) ​

## 说明 ​

可读写

返回指定RecordRange的字段Id，如果RecordRange未指定数据，则返回undefined

## 返回值 ​

Array

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
   const app = instance.Application;
   await instance.Application.ActiveView.Selection.FieldId
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   Application.ActiveView.Selection.FieldId
}
main()
```
 
# 297 API文档 / API / RecordRange / Font

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置单元格的字体颜色 ​

RecordRange.Font

## 说明 ​

可读写

返回当前RecordRange首个单元格的字体属性[Font](/documents/app-integration-dev/guide/dbsheet/Api/Font.html)

## 返回值 ​

[Font](/documents/app-integration-dev/guide/dbsheet/Api/Font.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const recordRange = await app.Sheet(1).RecordRange(1);
    const font = await recordRange.Font
    font.Color = "#ff00ff"
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const recordRange = Application.Sheet(1).RecordRange(1);
    const font = recordRange.Font
    font.Color = "#ff00ff"
}
main()
```
 
# 298 API文档 / API / RecordRange / Id

本页内容

# RecordRange.Id(属性) ​

## 说明 ​

可读写

返回指定RecordRange的记录id，如果RecordRange未指定数据，则返回undefined

## 返回值 ​

Array

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
   const app = instance.Application;
   await instance.Application.ActiveView.RecordRange(1).Id
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   Application.ActiveView.RecordRange(1).Id
}
main()
```
 
# 299 API文档 / API / RecordRange / Index

本页内容

# RecordRange.Index(属性) ​

## 说明 ​

可读写

返回指定RecordRange的记录索引，如果RecordRange未指定数据，则返回undefined

## 返回值 ​

Array

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
   const app = instance.Application;
   const index = await instance.Application.ActiveView.RecordRange("a").Index
   if (index && index.length > 0) {
        console.log(index)
   } else {
        console.log("记录不存在")
   }
   
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const index = Application.ActiveView.RecordRange("a").Index
   if (index && index.length > 0) {
        console.log(index)
   } else {
        console.log("记录不存在")
   }
}
main()
```
 
# 300 API文档 / API / RecordRange / Interior

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置单元格的填充颜色 ​

RecordRange.Interior

## 说明 ​

可读 返回当前RecordRange首个单元格的填充属性[Interior](/documents/app-integration-dev/guide/dbsheet/Api/Interior.html)

## 返回值 ​

[Interior](/documents/app-integration-dev/guide/dbsheet/Api/Interior.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const recordRange = await app.Sheet(1).RecordRange(1);
    const Interior = await recordRange.Interior
    Interior.Color = "#ff00ff"
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const recordRange = Application.Sheet(1).RecordRange(1);
    const Interior = recordRange.Interior
    Interior.Color = "#ff00ff"
}
main()
```
 
# 301 API文档 / API / RecordRange / Text

本页内容

# RecordRange.Text(属性) ​

## 说明 ​

只读 返回指定区域的文本值

## 返回值 ​

String/Array

这里返回的文本根据区域会有三种情况,RecordRange 选择器参看[RecordRange](/documents/app-integration-dev/guide/dbsheet/Api/RecordRange.html)的说明 1、指定的区域为一条记录一个字段，返回值为 字符串

```javascript
Application.ActiveSheet.RecordRange(1,1).Text
```

2、指定的区域为一条记录的多个字段，返回值为二维数组

```javascript
Application.ActiveSheet.RecordRange(1,[1,2]).Text
```

3、指定的区域为多条记录的多个字段，返回值为二维数组

```javascript
Application.ActiveSheet.RecordRange([1,2],[1,2]).Text
```

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
   const app = instance.Application;
   console.log(await app.ActiveSheet.RecordRange([1,2],[1,2]).Text)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  console.log(Application.ActiveSheet.RecordRange([1,2],[1,2]).Text)
}
main()
```
 
# 302 API文档 / API / RecordRange / Value

本页内容

- 说明
- 返回值
    - 这里返回的文本根据区域会有三种情况,RecordRange 选择器参看RecordRange的说明
    - 设置值的时候，也有三种设置值的方式
    - 不同字段类型设置Value时的数据结构
    - 部分不支持设置值的字段类型
- 浏览器环境示例
- 脚本编辑器 示例

# 设置单元格内容 ​

RecordRange.Value

## 说明 ​

可读写

RecordRange的值，读取和设置记录数据

## 返回值 ​

DbCellValue

### 这里返回的文本根据区域会有三种情况,RecordRange 选择器参看[RecordRange](/documents/app-integration-dev/guide/dbsheet/Api/RecordRange.html)的说明 ​

1、指定的区域为一条记录一个字段，返回单个值（值的类型看具体字段）

```javascript
Application.ActiveSheet.RecordRange(1,1).Value
```

2、指定的区域为一条记录的多个字段，返回值为二维数组

```javascript
Application.ActiveSheet.RecordRange(1,[1,2]).Value
```

3、指定的区域为多条记录的多个字段，返回值为二维数组

```javascript
Application.ActiveSheet.RecordRange([1,2],[1,2]).Value
```

### 设置值的时候，也有三种设置值的方式 ​

1、设置为单个值,将这个值设置到整个区域的指定单元格

```javascript
Application.ActiveSheet.RecordRange([1,2],[1,2]).Value = "1"
```

2、传入一维数组，目标是一条记录，则将数组设置到目标区域，如果目标是多条记录，则会将相同的数据设置到所有记录。

```javascript
Application.ActiveView.RecordRange([1,2],[1,2]).Value = ["1","2"]
```

3、传入二维数组，如果二维数组的长度为1，目标是一条记录，则将二维数组[0]设置到目标区域，如果目标是多条记录，则会将相同的数据设置到所有记录。如果二维数组为 M x N，则按顺序传入到目标区域。

```javascript
Application.ActiveView.RecordRange([1,2],[1,2]).Value = [["1","2"],["3","4"]]

// 数组长度为1时等价于 [["1","2"],["1","2"]]
Application.ActiveView.RecordRange([1,2],[1,2]).Value = [["1","2"]]
```

### 不同字段类型设置Value时的数据结构 ​

地址字段类型: 通过DBCellValue() 生成字段的数据

```javascript
Application.Sheets(1).Views(2).RecordRange(1, "@地址").Value = DBCellValue({districts:["广东省","珠海市","香洲区"],detail:"前岛环路xxxx号"})
```

级联字段类型：通过DBCellValue() 生成字段的数据

```javascript
Application.Sheets(1).Views(2).RecordRange(1, "@级联选项").Value = DBCellValue({districts:["广东省","珠海市","香洲区"]})
```

超链接字段类型：

```javascript
Application.RecordRange(1, "@超链接").Value = Application.DBCellValue({address:"wps.cn", display:"wps"})
```

关联字段类型: 参数传入关联的记录id

```javascript
const DBCellValue = Application.DBCellValue
Application.Sheets(1).Views(2).RecordRange(1, "@关联：数据表").Value = DBCellValue(["b","V"])
```

多选项类型：

```javascript
Application.Sheets(1).Views(2).RecordRange([5,6], 4).Value =[[DBCellValue(["未开始","进行中"])], DBCellValue(["进行中"])]
```

联系人字段类型：直接传入联系人的id，如果有多个联系人，可用","进行分割

```javascript
Application.Sheets(1).Views(2).RecordRange([5,6], "@联系人").Value = "238777563"
```

图片与附件字段：可以传入包含 URL/base64编码的图片/云文档 的数组，支持多个附件。 注意：由于脚本有运行时长限制，附件较大/或者较多时会导致超时，设置失败

```javascript
Application.Sheets(1).Views(2).RecordRange(1, "@图片和附件").Value = DBCellValue([{fileData: url/base64, fileName: ""}])
```

如果要在原来的附件上，新增其它附件可以先读出来的数组增加后再重新设置

```javascript
const range = Application.Sheets(1).Views(2).RecordRange(1, "@图片和附件")
const dbCellValue = range.Value
const attments = dbCellValue.Value
attments.push({fileData: url/base64, fileName: ""})
range.Value = oldValue
```

其它字段类型可以直接使用字符串，日期和时间类型必须符合日期的格式的字符串

### 部分不支持设置值的字段类型 ​

```javascript
DbSheetFieldType.Formula, // 公式字段
DbSheetFieldType.Lookup, // 引用字段
DbSheetFieldType.CreatedBy, // 创建者字段
DbSheetFieldType.Note, // 富文本字段
DbSheetFieldType.SearchLookup, // 查找引用字段
DbSheetFieldType.Button,  // 按钮字段 
DbSheetFieldType.LastModifiedBy, // 最近修改者字段
DbSheetFieldType.CreatedTime, // 创建时间字段
DbSheetFieldType.LastModifiedTime, //最后修改时间字段
DbSheetFieldType.AutoNumber, // 自动编号 
DbSheetFieldType.Automations, // 自动任务
```

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
   const app = instance.Application;
   app.Sheets(1).Views(2).RecordRange([5,6], 1).Value = "1111111"
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   Application.Sheets(1).Views(2).RecordRange([5,6], 1).Value = "1111111"
}
main()
```
 
# 303 API文档 / API / RecordRange / Add

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 插入记录 ​

## 说明 ​

插入新的记录，在指定行记录之前或之后插入

## 语法 ​

表达式.Add(Before,After,Count)

表达式:RecordRange

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Before | 否 | Number/String | 在行记录之前添加，Number时指定插入时索引，String时指定插入ID |
| After | 否 | Number/String | 在行记录之后添加，Number时指定插入时索引，String时指定插入ID |
| Count | 否 | Number | 一次插入N条数据，未指定时插入1条 |

## 返回值 ​

Self

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  // 在第1行，向上方添加10条记录
  const range = await app.ActiveView.RecordRange.Add(1, null, 10)
  
  const template = ["商品"]
  const range1 = []
  // 给1-10行赋值
  for (let i = 1; i < 11; i++ ) {
    range1.push([template[0]+i,i])
  }
  range.Value = range1

}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const range = Application.ActiveView.RecordRange.Add(31, null, 5)
  // 将插入的5条记录的名称字段 初始化为 “名称”
  range.Item(undefined, "@名称").Value = "名称"
}
main()
```
 
# 304 API文档 / API / RecordRange / Condition

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 筛选记录 ​

## 说明 ​

筛选符合条件的记录

## 语法 ​

表达式.Condition(Filters,FilterOp)

表达式:RecordRange

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Filters | 是 | Filter[] | 筛选数据的条件组, 条件组里每个Filter，可以包含多个筛选条件 |
| FilterOp | 否 | "And"/"Or" | 筛选数据条件组之间的关系，是同时满足还是只需要满足一条，默认值为And |

Filter数据结构：

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Criterias | 是 | [Criteria](/documents/app-integration-dev/guide/dbsheet/Api/Criteria.html)[] | 筛选条件项组 |
| Op | 否 | "And"/"Or" | 筛选条件项组之间的关系，是同时满足还是只需要满足一条，默认值为And |

[Criteria](/documents/app-integration-dev/guide/dbsheet/Api/Criteria.html) 中筛选规则（大小写不敏感）：

| 枚举值 | 描述 |
| --- | --- |
| Equals | 等于 |
| NotEqu | 不等于 |
| Greater | 大于 |
| GreaterEqu | 大等于 |
| Less | 小于 |
| LessEqu | 小等于 |
| GreaterEquAndLessEqu | 介于（取等） |
| LessOrGreater | 不介于（不取等） |
| BeginWith | 开头是 |
| EndWith | 结尾是 |
| Contains | 包含 |
| NotContains | 不包含 |
| Intersected | 指定值 |
| Empty | 为空 |
| NotEmpty | 不为空 |

各筛选规则独立地限制了values数组内最多允许填写的元素数，当values内元素数超过阈值时，该筛选规则将失效。“为空、不为空”不允许填写元素；“介于”允许最多填写2个元素；“指定值”允许填写65535个元素；其他规则允许最多填写1个元素 values[]数组内的元素为字符串时，表示文本匹配。目前还支持对日期进行动态筛选，此时values[]内的元素需以结构体的形式给出：

```javascript
const dateValue = {"dynamicType": "lastMonth","type": "DynamicSimple"}
Criteria("@日期", "Equals", [dateValue])
```

上述示例对应的筛选条件为“等于上一个月”。 要使用日期动态筛选，values[]内的结构体需要指定"type": "DynamicSimple"，当"op"为"equals"时，"dynamicType"可以为如下的值（大小写不敏感）：

| 枚举值 | 描述 |
| --- | --- |
| today | 今天 |
| yesterday | 昨天 |
| tomorrow | 明天 |
| last7Days | 最近7天 |
| last30Days | 最近30天 |
| thisWeek | 本周 |
| lastWeek | 上周 |
| nextWeek | 下周 |
| thisMonth | 本月 |
| lastMonth | 上月 |
| nextMonth | 次月 |

当"op"为"greater"或"less"时，"dynamicType"只能是昨天、今天或明天。

对不同字段类型，values会有不同的用法 联系人字段:

```javascript
// value是一个结构体，指定type为 Contact, value 为用户id
const dateValue = {"type":"Contact", value:"user id"}
```

单/多选项字段:

```javascript
// value是一个结构体，指定type为 SelectItem, value 为选项的ID
const dateValue = {"type":"SelectItem", value:"B"}
```

## 返回值 ​

RecordRange

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  // 创建筛选条件criteria
  const Criteria = app.Criteria 
  const criterias = []
  criterias.push(await Criteria("@名称",  "Intersected", ["1"]))
  // 创建筛选列表filters
  const filters = []
  const filter = {Criterias: criterias, Op: "AND"}
  filters.push(filter)
  // 筛选记录
  const res = await app.ActiveSheet.Views(1).RecordRange.Condition(filters, "AND")
  console.log(res)
  // 操作记录，返回第一个筛选结果的记录ID
  await res.Item(1).Id
  // 删除记录
  await res.Item(1).Delete()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  // 创建筛选条件criteria
    const critera1 = Criteria("@名称", "Intersected", ["1", "999", "aaaa"])
    const critera2 = Criteria("@数量", "Equals", ["1"])
    const criterias = []
    criterias.push(critera1)
    criterias.push(critera2)
    // 创建filters
    const filters = []
    const filter = { Criterias: criterias, Op: "OR" }
    filters.push(filter)
    const res = Application.ActiveSheet.Views(1).RecordRange.Condition(filters, "AND")
    console.log(res.Value)
}
main()
```
 
# 305 API文档 / API / RecordRange / Delete

本页内容

# RecordRange.Delete(方法) ​

## 说明 ​

删除记录

## 语法 ​

表达式.Delete()

表达式:RecordRange

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |

## 返回值 ​

ApiResult

## jsApi 示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  await instance.Application.ActiveView.RecordRange(41).Delete()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  Application.ActiveView.RecordRange(41).Delete()
}
main()
```
 
# 306 API文档 / API / RecordRange / Item

本页内容

# RecordRange.Item(方法) ​

## 说明 ​

获取指定索引位置的记录 1、通过数值指定索引的记录

```javascript
RecordRange.Item(1)
```

2、通过指定记录ID来指定记录

```javascript
RecordRange.Item("a")
```

3、通过符号:可以指定索引范围的记录

```javascript
RecordRange.Item("1:100")
```

4、通过数组可以指定不连续的记录

```javascript
RecordRange.Item(["1","5","10:20"])
```

5、通过参数2为字符串指定字段ID可以限定显示的字段

```javascript
RecordRange.Item(1, "a")
```

6、通过参数2为数组可以指定多个字段ID可以限定显示的字段

```javascript
RecordRange.Item(1, ["a","b","A"])
```

7、通过参数2如果以@字符作为首字符表示字段名称

```javascript
RecordRange.Item(1, ["@名称", "@数量"])
```

## 语法 ​

表达式.Item(Index,Field)

表达式:RecordRange

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Index | 否 | number/string/array | 传入number时索引从1开始，传入字符串时表示id |
| Field | 否 | number/string/array | 传入number时索引从1开始，传入字符串时如果字符串以@开始则表示为字段名称，否则为字段id |

## 返回值 ​

Self

## jsApi 示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const values = await app.ActiveView.RecordRange.Item(1, ["@名称", "@日期", "@数量"]).Value
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const values = Application.ActiveView.RecordRange.Item(1, ["@名称", "@日期", "@数量"]).Value
}
main()```
```
 
# 307 API文档 / API / RecordRange / Select

本页内容

# RecordRange.Select(方法) ​

## 说明 ​

插入新的记录

## 语法 ​

表达式.Select()

表达式:RecordRange

## 参数 ​

无

## 返回值 ​

Boolean

## jsApi 示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const range = await app.ActiveView.RecordRange(1)
  range.Select()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const range = Application.ActiveView.RecordRange(1)
  // 将插入的5条记录的名称字段 初始化为 “名称”
  range.Select()
}
main()
```
 
# 308 API文档 / API / RecordRange / SetValues

本页内容

# RecordRange.SetValues(方法) ​

## 说明 ​

设置指定单元格的值

## 语法 ​

表达式.SetValues(Values, ignoreErr)

表达式:RecordRange

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Values | 是 | Array | 设置到单元格的值 |
| ignoreErr | 否 | bool | 默认值为flase,当设置为true时，输入数据中某一条数据出错，这条数据不写入，其它数据正常写入；当设置为false时，其中一条数据出错，则都不写入 |

## 返回值 ​

ApiResult

## jsApi 示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const Application = instance.Application;
  const res = await WPSOpenApi.Application.RecordRange("1:2", ["@名称", "@数量"]).SetValues([["1111", 2222]], true)
  if (res.code !== 0) {
    console.log(res.Message)
  }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {

  const res = Application.RecordRange("1:2", ["@名称", "@数量"]).SetValues([["1111", 2222]], true)
  if (res.code !== 0) {
    console.log(res.Message)
  }
}
main()
```
 
# 309 API文档 / API / RecordRange / OnDeleteRecord

本页内容

- 说明
- 语法
- 参数
- 返回值
- 事件返回数据
- 事件返回数据示例
- 浏览器环境示例
- 脚本编辑器 示例

# 监听删除记录的事件 ​

RecordRange.OnDeleteRecord(方法)

## 说明 ​

为 RecordRange 添加 DeleteRecord 事件,当删除 RecordRange 时触发。注意在脚本编辑器中使用时，脚本运行结束就会退出运行，这时可能回调无法被正常触发

## 语法 ​

表达式: OnDeleteRecord(Callback)

表达式: RecordRange

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Callback | 是 | func | 指定事件发生时的回调方法 ; const eventContext = await RecordRange.OnDeleteRecord(()=&gt;{ ... }) |

## 返回值 ​

EventContext

## 事件返回数据 ​

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| sheetId | Number | 表的 Id |
| recordIds | Array | 记录集合的 Ids |

## 事件返回数据示例 ​

```javascript
{
    recordIds: ['A','C']
    sheetId: 1
}
```

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    let eventContext;
    eventContext = await app
        .Sheets(1)
        .Views(1)
        .RecordRange(1)
        .OnDeleteRecord(data => {
            console.log(data);
            // 取消事件监听
            eventContext.Destroy();
        });
    app.Sheets(1).Views(1).RecordRange(1).Delete();
    //这里会执行OnDeleteRecord的回调
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    let eventContext;
    eventContext = Application.Sheets(1)
        .Views(1)
        .RecordRange(1)
        .OnDeleteRecord(data => {
            console.log(data);
            // 取消事件监听
            eventContext.Destroy();
        });
    Application.Sheets(1).Views(1).RecordRange(1).Delete();
    //这里会执行OnDeleteRecord的回调
}
main();
```
 
# 310 API文档 / API / RecordRange / OnUpdate

本页内容

- 说明
- 语法
- 参数
- 返回值
- 事件返回数据
- 浏览器环境示例
- 脚本编辑器 示例

# 监听修改记录的事件 ​

RecordRange.OnUpdate(方法)

## 说明 ​

为 RecordRange 添加 Update 事件,当更新 RecordRange 时触发。注意在脚本编辑器中使用时，脚本运行结束就会退出运行，这时可能回调无法被正常触发

## 语法 ​

表达式: OnUpdate(Callback)

表达式: RecordRange

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Callback | 是 | func | 指定事件发生时的回调方法 ; const eventContext = await RecordRange.OnUpdate(()=&gt;{ ... }) |

## 返回值 ​

EventContext

## 事件返回数据 ​

[RecordRange]

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    let eventContext;
    eventContext = await app
        .Sheets(1)
        .Views(1)
        .RecordRange(1)
        .OnUpdate(data => {
            console.log(data);
            // 取消事件监听
            eventContext.Destroy();
        });
    app.Sheets(1).Views(1).RecordRange(1).Value = ['名称111', 4, '选项1'];
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    let eventContext;
    eventContext = Application.Sheets(1)
        .Views(1)
        .RecordRange(1)
        .OnUpdate(data => {
            console.log(data);
            // 取消事件监听
            eventContext.Destroy();
        });
    Application.Sheets(1).Views(1).RecordRange(1).Value = ['名称111', 4, '选项1'];
}
main();
```
 
# 311 API文档 / API / Records / Records对象

本页内容

# Records (对象) ​

## 说明 ​

Records 对象，表示记录相关集合，用户可通过Records对象直接访问记录和字段 支持以下几种访问方式 View.Records.Item(记录id) View.Records.Item(记录索引) View.Records(记录ID/记录索引, 字段索引/字段名称) View.Records(记录ID/记录索引).Item(字段索引/字段名称)

## 方法 ​

- [Item](/documents/app-integration-dev/guide/dbsheet/Api/Records_Item.html)
- [Add](/documents/app-integration-dev/guide/dbsheet/Api/Records_Add.html)
- [Delete](/documents/app-integration-dev/guide/dbsheet/Api/Records_Delete.html)
- [FindNext](/documents/app-integration-dev/guide/dbsheet/Api/Records_FindNext.html)
- [FindPrevious](/documents/app-integration-dev/guide/dbsheet/Api/Records_FindPrevious.html)

## 属性 ​

- [Count](/documents/app-integration-dev/guide/dbsheet/Api/Groups_Count.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const record = await app.Sheets(1).Views(1).Records(10)
    // 返回字段的文本表示
    console.log(await record.Text)
    // 返回字段的值
    console.log(await record.Value)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const record = Application.Sheets(1).Views(1).Records(10)
    // 返回字段的文本表示
    console.log(record.Text)
    // 返回字段的值
    console.log(record.Value)
 }
main()
```
 
# 312 API文档 / API / Records / Count

本页内容

# Records.Count(属性) ​

## 说明 ​

可读 返回Records列表的数量

## 返回值 ​

Number

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const records = await app.Sheets(1).Views(1).Records
    // 返回记录数量
    console.log(await records.Count)
    const insertRecords = await app.Sheets(1).Views(1).Records.Add(10, undefined, 2) // 在第10行，向上方添加2条记录
    // 输出 2
    console.log(await insertRecords.Count)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const records = Application.Sheets(1).Views(1).Records
    // 返回记录数量
    console.log(records.Count)
    const insertRecords = Application.Sheets(1).Views(1).Records.Add(10, undefined, 2) // 在第10行，向上方添加2条记录
    // 输出 2
    console.log(insertRecords.Count)
 }
main()
```
 
# 313 API文档 / API / Records / Add

本页内容

# 添加记录 ​

## 说明 ​

添加一条或者多条记录

## 语法 ​

表达式：Add(Before, After, Count)

表达式：Records

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Before | 否 | number/string | 指定的记录对象（索引从1开始/记录ID），新建的记录将基于此之前 |
| After | 否 | number/string | 指定的记录对象（索引从1开始/记录ID），新建的记录将基于此之后 |
| Count | 否 | number | 插入多少条记录，默认值为1条 |

## 返回值 ​

ApiResult

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    app.Sheets(1).Views(1).Records.Add(undefined, 10, 2) // 在第10行，向下方添加2条记录
    const records = await app.Sheets(1).Views(1).Records.Add(10, undefined, 2) // 在第10行，向上方添加2条记录
    // 在插入的第一条记录的首字段填入 "123"
    records.Item(1, 1).Value = "123"
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const records = Application.Sheets(1).Views(1).Records.Add(10, undefined, 2) // 在第10行，向上方添加2条记录
    // 在插入的第一条记录的首字段填入 "123"
    records.Item(1, 1).Value = "123"
 }
main()
```
 
# 314 API文档 / API / Records / Delete

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 删除记录 ​

## 说明 ​

删除某行记录

## 语法 ​

表达式：Delete(Index)

表达式：Records

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Index | 是 | number/string | 索引从1开始/记录ID |

## 返回值 ​

ApiResult

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    // 删除视图上的第一条数据
    app.Sheets(1).Views(1).Records.Delete(1)
    // 在第10行，向上方添加2条记录
    const records = await app.Sheets(1).Views(1).Records.Add(10, undefined, 2) 
    // 删除刚插入的第一条数据
    records.Delete(1)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    // 删除视图上的第一条数据
    Application.Sheets(1).Views(1).Records.Delete(1)
    // 在第10行，向上方添加2条记录
    const records = Application.Sheets(1).Views(1).Records.Add(10, undefined, 2) 
    // 删除刚插入的第一条数据
    records.Delete(1)
 }
main()
```
 
# 315 API文档 / API / Records / FindNext

本页内容

# Application.FindNext (方法) ​

## 说明 ​

1. 查找匹配相同条件（当前活动单元格）的下一个单元格，并返回表示该单元格的 Record 对象，该操作不影响选定内容和活动单元格。
2. 持续地调用该方法会返回同一个查找单元格，需要结合`Record.select()`方法移动活动单元格后再调用，才会继续查找下一个值。

## 语法 ​

表达式: FindNext(What)

表达式: Records

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| What | 是 | String | 查找值 |

## 返回值 ​

Record | Null

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const record = await app.Sheets(1).Views(1).Records.FindNext('查找值');
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const record = Application.Sheets(1).Views(1).Records.FindNext('查找值');
}
main()
```
 
# 316 API文档 / API / Records / FindPrevious

本页内容

# Application.FindPrevious (方法) ​

## 说明 ​

1. 查找匹配相同条件（当前活动单元格）的上一个单元格，并返回表示该单元格的 Record 对象，该操作不影响选定内容和活动单元格。
2. 持续地调用该方法会返回同一个查找单元格，需要结合`Record.select()`方法移动活动单元格后再调用，才会继续查找上一个值。

## 语法 ​

表达式: FindPrevious(What)

表达式: Records

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| What | Y | String | 查找值 |

## 返回值 ​

Record | Null

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const record = await app.Sheets(1).Views(1).Records.FindPrevious('查找值');
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const record = Application.Sheets(1).Views(1).Records.FindPrevious('查找值');
}
main()
```
 
# 317 API文档 / API / Records / Item

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 查看记录 ​

## 说明 ​

获取指定索引行的记录信息

## 语法 ​

表达式：Item(Index)

表达式：Records

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Index | 是 | number/string | 索引从1开始/记录ID |

## 返回值 ​

Record

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const record = await app.Sheets(1).Views(1).Records.Item(10)
    // 返回字段的文本表示
    console.log(await record.Text)
    // 返回字段的值
    console.log(await record.Value)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const record = Application.Sheets(1).Views(1).Records.Item(10)
    // 返回字段的文本表示
    console.log(record.Text)
    // 返回字段的值
    console.log(record.Value)
 }
main()
```
 
# 318 API文档 / API / AddressField / AddressField对象

本页内容

# AddressField (对象) ​

## 说明 ​

AddressField 地址字段的属性，修改属性之后需要 FieldDescriptor 调用Apply()方法使得修改生效。如果非地址字段，无法设置相关属性

## 属性 ​

- [IsDetailedAddress](/documents/app-integration-dev/guide/dbsheet/Api/AddressField_IsDetailedAddress.html)
- [IsUsePresetAddress](/documents/app-integration-dev/guide/dbsheet/Api/AddressField_IsUsePresetAddress.html)
- [Level](/documents/app-integration-dev/guide/dbsheet/Api/AddressField_Level.html)
- [PresetAddress](/documents/app-integration-dev/guide/dbsheet/Api/AddressField_PresetAddress.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const fieldDescriptor = await app.Sheets("数据表").FieldDescriptors("@地址")
    const prop = await fieldDescriptor.Address
    prop.Text = "Get it"
    fieldDescriptor.Apply()
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@地址")
    const prop = fieldDescriptor.Address
    prop.Text = "Get it"
    fieldDescriptor.Apply()
 }
main()
```
 
# 319 API文档 / API / AddressField / IsDetailedAddress

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 填写详细地址 ​

AddressField.IsDetailedAddress(属性)

## 说明 ​

可读写

地址字段是否需要填写详细地址

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const Application = instance.Application;
   const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@地址")
   const prop = await fieldDescriptor.Address
   prop.IsDetailedAddress = false
   fieldDescriptor.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@地址")
   const prop = fieldDescriptor.Address
   prop.IsDetailedAddress = false
   fieldDescriptor.Apply()
}
main()
```
 
# 320 API文档 / API / AddressField / IsUsePresetAddress

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 预设指定地址 ​

AddressField.IsUsePresetAddress(属性)

## 说明 ​

可读写

地址字段 是否预设指定地址,当设置了IsUsePresetAddress为false后，PresetAddress不生效

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const Application = instance.Application;
   const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@地址")
   const prop = await fieldDescriptor.Address
   prop.IsUsePresetAddress = false
   fieldDescriptor.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@地址")
   const prop = fieldDescriptor.Address
   prop.IsUsePresetAddress = false
   fieldDescriptor.Apply()
}
main()
```
 
# 321 API文档 / API / AddressField / Level

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 地址级别数 ​

AddressField.Level(属性)

## 说明 ​

可读写

地址字段格式包含几级地址，1-5级

| 级数 | 示例 |
| --- | --- |
| 1 | 省 - 详细地址 |
| 2 | 省/市 - 详细地址 |
| 3 | 省/市/区 - 详细地址 |
| 4 | 省/市/区/街道 - 详细地址 |
| 5 | 省/市/区/街道/社区 - 详细地址 |

## 返回值 ​

Number

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const Application = instance.Application;
   const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@地址")
   const prop = await fieldDescriptor.Address
   prop.Level = 5
   fieldDescriptor.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@地址")
   const prop = fieldDescriptor.Address
   prop.Level = 5
   fieldDescriptor.Apply()
}
main()
```
 
# 322 API文档 / API / AddressField / PresetAddress

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 默认值 ​

AddressField.PresetAddress(属性)

## 说明 ​

可读写

地址字段的默认值,当设置了IsUsePresetAddress为false后，PresetAddress不生效

## 返回值 ​

Object

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const Application = instance.Application;
   const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@地址")
   const prop = await fieldDescriptor.Address
   prop.IsUsePresetAddress = true
   prop.PresetAddress = {detail:"", districts:["广东省","江门市"]}
   fieldDescriptor.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@地址")
   const prop = fieldDescriptor.Address
   prop.IsUsePresetAddress = true
   prop.PresetAddress = {detail:"", districts:["广东省","江门市"]}
   fieldDescriptor.Apply()
}
main()
```
 
# 323 API文档 / API / Attachment / Attachment对象

本页内容

# Attachment (对象) ​

## 说明 ​

图片和附件字段返回的数据结构，可以通过里面的属性读取文件的信息和地址

## 方法 ​

## 属性 ​

- [FileId](/documents/app-integration-dev/guide/dbsheet/Api/Attachment_FileId.html)
- [FileName](/documents/app-integration-dev/guide/dbsheet/Api/Attachment_FileName.html)
- [FileSize](/documents/app-integration-dev/guide/dbsheet/Api/Attachment_FileSize.html)
- [FileType](/documents/app-integration-dev/guide/dbsheet/Api/Attachment_FileType.html)
- [ImgSize](/documents/app-integration-dev/guide/dbsheet/Api/Attachment_ImgSize.html)
- [LinkUrl](/documents/app-integration-dev/guide/dbsheet/Api/Attachment_LinkUrl.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const dbCellValue = await app.Sheets(1).Views(1).RecordRange(1,"@图片和附件").Value
 const attments = await dbCellValue.Value
 console.log(await attments[0].FileName)
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
 const dbCellValue = Application.Sheets(1).Views(1).RecordRange(1,"@图片和附件").Value
 const attments = dbCellValue.Value
 console.log(attments[0].FileName)
 }
main()
```
 
# 324 API文档 / API / Attachment / FileId

本页内容

# Attachment.FileId(属性) ​

## 说明 ​

只读 返回附件字段的Id

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const app = instance.Application;
   const value = await app.Sheets(1).Views(1).RecordRange(1,"@图片和附件").Value
   const attments = await value.Value
   console.log(await attments[0].FileId)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const value = ActiveView.RecordRange(1, "@图片和附件").Value
  const attments = value.Value
  console.log(attments[0].FileId)
}
main()
```
 
# 325 API文档 / API / Attachment / FileName

本页内容

# Attachment.FileName(属性) ​

## 说明 ​

只读 返回附件名称

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
   const app = instance.Application;
   const value = await app.Sheets(1).Views(1).RecordRange(1,"@图片和附件").Value
   const attments = await value.Value
   console.log(await attments[0].FileName)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main(){
  const value = ActiveView.RecordRange(1, "@图片和附件").Value
  const attments = value.Value
  console.log(attments[0].FileName)
}
main()
```
 
# 326 API文档 / API / Attachment / FileSize

本页内容

# Attachment.FileSize(属性) ​

## 说明 ​

只读 返回附件大小

## 返回值 ​

Number

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const app = instance.Application;
   const value = await app.Sheets(1).Views(1).RecordRange(1,"@图片和附件").Value
   const attments = await value.Value
   console.log(await attments[0].FileSize)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const value = ActiveView.RecordRange(1, "@图片和附件").Value
  const attments = value.Value
  console.log(attments[0].FileSize)
}
main()
```
 
# 327 API文档 / API / Attachment / FileType

本页内容

# Attachment.FileType(属性) ​

## 说明 ​

只读 返回附件类型

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const app = instance.Application;
   const value = await app.Sheets(1).Views(1).RecordRange(1,"@图片和附件").Value
   const attments = await value.Value
   console.log(await attments[0].FileType)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const value = ActiveView.RecordRange(1, "@图片和附件").Value
  const attments = value.Value
  console.log(attments[0].FileType)
}
main()
```
 
# 328 API文档 / API / Attachment / ImgSize

本页内容

# Attachment.ImgSize(属性) ​

## 说明 ​

只读 返回图片的尺寸

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const app = instance.Application;
   const value = await app.Sheets(1).Views(1).RecordRange(1,"@图片和附件").Value
   const attments = await value.Value
   console.log(await attments[0].ImgSize)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const value = ActiveView.RecordRange(1, "@图片和附件").Value
  const attments = value.Value
  console.log(attments[0].ImgSize)
}
main()
```
 
# 329 API文档 / API / Attachment / LinkUrl

本页内容

# Attachment.LinkUrl(属性) ​

## 说明 ​

只读 返回附件的链接地址，如果是云文档则返回云文档的地址，如果是图片则返回图片的下载地址

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const app = instance.Application;
   const value = await app.Sheets(1).Views(1).RecordRange(1,"@图片和附件").Value
   const attments = await value.Value
   console.log(await attments[0].LinkUrl)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const value = ActiveView.RecordRange(1, "@图片和附件").Value
  const attments = value.Value
  console.log(attments[0].LinkUrl)
}
main()
```
 
# 330 API文档 / API / AttachmentField / AttachmentField对象

本页内容

# AttachmentField (对象) ​

## 说明 ​

AttachmentField 图片和附件字段的属性，修改属性之后需要 FieldDescriptor 调用Apply()方法使得修改生效。如果非图片和附件字段，无法设置相关属性

## 属性 ​

- [DisplayStyle](/documents/app-integration-dev/guide/dbsheet/Api/AttachmentField_DisplayStyle.html)
- [IsOnlyCameraUpload](/documents/app-integration-dev/guide/dbsheet/Api/AttachmentField_IsOnlyCameraUpload.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const field = await app.Sheets(1).FieldDescriptors("@图片和附件")
  const prop = await field.Attachment
  prop.DisplayStyle = "Pic"
  prop.IsOnlyCameraUpload = false
  field.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main(){
  const field = Application.Sheets(1).FieldDescriptors("@图片和附件")
  const prop = field.Attachment
  prop.DisplayStyle = "Pic" 
  prop.IsOnlyCameraUpload = true
  field.Apply()
}
main()
```
 
# 331 API文档 / API / AttachmentField / DisplayStyle

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置显示样式 ​

AttachmentField.DisplayStyle(属性)

## 说明 ​

可读写

图片和附件字段的显示样式，是以缩略图样式显示还是以列表的样式显示

## 返回值 ​

Enum.DbAttachmentDisplayStyle

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const field = await app.Sheets(1).FieldDescriptors("@图片和附件")
  const prop = await field.Attachment
  prop.DisplayStyle = "Pic"
  field.Apply()
  console.log(await prop.DisplayStyle)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const field = Application.Sheets(1).FieldDescriptors("@图片和附件")
  const prop = field.Attachment
  prop.DisplayStyle = "List"
  field.Apply()
}
main()
```
 
# 332 API文档 / API / AttachmentField / IsOnlyCameraUpload

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 仅可通过移动端拍摄上传 ​

AttachmentField.IsOnlyCameraUpload(属性)

## 说明 ​

可读写

图片和附件字段是否仅可通过移动端拍摄上传

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const field = await app.Sheets(1).FieldDescriptors("@图片和附件")
  const prop = await field.Attachment
  prop.IsOnlyCameraUpload = true
  field.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const field = Application.Sheets(1).FieldDescriptors("@图片和附件")
  const prop = field.Attachment
  prop.IsOnlyCameraUpload = true
  field.Apply()
}
main()
```
 
# 333 API文档 / API / AutoLinkCondition / AutoLinkCondition对象

本页内容

# AutoLinkCondition (对象) ​

## 说明 ​

单向关联和双向关联字段的自动匹配条件

## 方法 ​

## 属性 ​

- [LinkSheetFieldId](/documents/app-integration-dev/guide/dbsheet/Api/AutoLinkCondition_LinkSheetFieldId.html)
- [SheetCondContents](/documents/app-integration-dev/guide/dbsheet/Api/AutoLinkCondition_SheetCondContents.html)
- [SheetCondType](/documents/app-integration-dev/guide/dbsheet/Api/AutoLinkCondition_SheetCondType.html)
- [OpType](/documents/app-integration-dev/guide/dbsheet/Api/AutoLinkCondition_OpType.html)
- [IntersectedConds](/documents/app-integration-dev/guide/dbsheet/Api/AutoLinkCondition_IntersectedConds.html)
- [DateIntersectedValues](/documents/app-integration-dev/guide/dbsheet/Api/AutoLinkCondition_DateIntersectedValues.html)

## 浏览器环境示例 ​

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const group = await app.Sheets(1).FieldDescriptors(2).AutoLinkGroups(1)
  const condition = group.Conditions(1)
  console.log(condition.LinkSheetFieldId) 
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const group = Application.Sheets(1).FieldDescriptors(2).AutoLinkGroups(1)
  const condition = group.Conditions(1)
  console.log(condition.LinkSheetFieldId) 
 }
main()
```
 
# 334 API文档 / API / AutoLinkCondition / DateIntersectedValues

本页内容

# AutoLinkCondition.DateIntersectedValues(属性) ​

## 说明 ​

可读写

自动匹配条件日期指定值的特殊参数

## 返回值 ​

Array

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const group = await app.Sheets(1).FieldDescriptors(2).AutoLinkGroups(1)
  const condition = group.Conditions(1)
  console.log(condition.DateIntersectedValues) 
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const group = Application.Sheets(1).FieldDescriptors(2).AutoLinkGroups(1)
  const condition = group.Conditions(1)
  console.log(condition.DateIntersectedValues)
 }
main()
```
 
# 335 API文档 / API / AutoLinkCondition / IntersectedConds

本页内容

# AutoLinkCondition.IntersectedConds(属性) ​

## 说明 ​

可读写

自动匹配条件，供“指定值”用的显示文本

## 返回值 ​

Array

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const group = await app.Sheets(1).FieldDescriptors(2).AutoLinkGroups(1)
  const condition = group.Conditions(1)
  console.log(condition.IntersectedConds) 
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const group = Application.Sheets(1).FieldDescriptors(2).AutoLinkGroups(1)
  const condition = group.Conditions(1)
  console.log(condition.IntersectedConds) 
 }
main()
```
 
# 336 API文档 / API / AutoLinkCondition / LinkSheetFieldId

本页内容

# AutoLinkCondition.LinkSheetFieldId(属性) ​

## 说明 ​

可读写

单向关联和双向关联字段的自动匹配条件，关联的数据表字段Id

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const group = await app.Sheets(1).FieldDescriptors(2).AutoLinkGroups(1)
  const condition = group.Conditions(1)
  console.log(condition.LinkSheetFieldId) 
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const group = Application.Sheets(1).FieldDescriptors(2).AutoLinkGroups(1)
  const condition = group.Conditions(1)
  console.log(condition.LinkSheetFieldId) 
 }
main()
```
 
# 337 API文档 / API / AutoLinkCondition / OpType

本页内容

# AutoLinkCondition.OpType(属性) ​

## 说明 ​

可读写

自动匹配条件的操作符

## 返回值 ​

DbFilterCriteriaOpType

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const group = await app.Sheets(1).FieldDescriptors(2).AutoLinkGroups(1)
  const condition = group.Conditions(1)
  console.log(condition.OpType) 
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const group = Application.Sheets(1).FieldDescriptors(2).AutoLinkGroups(1)
  const condition = group.Conditions(1)
  console.log(condition.OpType) 
 }
main()
```
 
# 338 API文档 / API / AutoLinkCondition / SheetCondContents

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# AutoLinkCondition.SheetCondContents(属性) ​

## 说明 ​

可读写

单向关联和双向关联字段的自动匹配条件

## 返回值 ​

Array

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const group = await app.Sheets(1).FieldDescriptors(2).AutoLinkGroups(1)
  const condition = group.Conditions(1)
  console.log(condition.SheetCondContents) 
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const group = Application.Sheets(1).FieldDescriptors(2).AutoLinkGroups(1)
  const condition = group.Conditions(1)
  console.log(condition.SheetCondContents) 
 }
main()
```
 
# 339 API文档 / API / AutoLinkCondition / SheetCondType

本页内容

# AutoLinkCondition.SheetCondType(属性) ​

## 说明 ​

可读写

单向关联和双向关联字段的自动匹配条件

## 返回值 ​

DbAutolinkCondType

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const group = await app.Sheets(1).FieldDescriptors(2).AutoLinkGroups(1)
  const condition = group.Conditions(1)
  console.log(condition.SheetCondType) 
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const group = Application.Sheets(1).FieldDescriptors(2).AutoLinkGroups(1)
  const condition = group.Conditions(1)
  console.log(condition.SheetCondType) 
 }
main()
```
 
# 340 API文档 / API / AutoLinkConditions / AutoLinkConditions对象

本页内容

# AutoLinkConditions (对象) ​

## 说明 ​

单向关联和双向关联字段的匹配条件集合

## 方法 ​

- [Item](/documents/app-integration-dev/guide/dbsheet/Api/AutoLinkConditions_Item.html)
- [Add](/documents/app-integration-dev/guide/dbsheet/Api/AutoLinkConditions_Add.html)
- [Delete](/documents/app-integration-dev/guide/dbsheet/Api/AutoLinkConditions_Delete.html)

## 属性 ​

- [Count](/documents/app-integration-dev/guide/dbsheet/Api/AutoLinkConditions_Count.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
  const app = instance.Application;
  const group = await app.Sheets(1).FieldDescriptors(2).AutoLinkGroups(1)
  const Conditions = await group.Conditions
  console.log(await Conditions.Count) 
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const group = Application.Sheets(1).FieldDescriptors(2).AutoLinkGroups(1)
  const Conditions = group.Conditions
  console.log(Conditions.Count)
 }
main()
```
 
# 341 API文档 / API / AutoLinkConditions / Count

本页内容

# AutoLinkConditions.Count(属性) ​

## 说明 ​

可读 单向关联和双向关联字段的匹配条件集合数

## 返回值 ​

Number

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const group = await app.Sheets(1).FieldDescriptors(2).AutoLinkGroups(1)
  const Conditions = await group.Conditions
  console.log(await Conditions.Count)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const group = Application.Sheets(1).FieldDescriptors(2).AutoLinkGroups(1)
  const Conditions = group.Conditions
  console.log(Conditions.Count)
}
main()
```
 
# 342 API文档 / API / AutoLinkConditions / Add

本页内容

# AutoLinkConditions.Add(方法) ​

## 说明 ​

添加单向关联和双向关联字段的自动匹配条件

## 语法 ​

表达式.Add(LinkSheetFieldId,SheetCondContents,SheetCondType,OpType,IntersectedConds,DateIntersectedValues)

表达式:AutoLinkConditions

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| LinkSheetFieldId | 是 | string | 引用表的字段ID |
| SheetCondContents | 否 | string |  |
| SheetCondType | 否 | DbAutolinkCondType | 关联类型 |
| OpType | 否 | DbFilterCriteriaOpType | 自动匹配条件的关系 |
| IntersectedConds | 否 |  | 供“指定值”用的显示文本 |
| DateIntersectedValues | 否 |  | 日期指定值的特殊参数 |

## 返回值 ​

AutoLinkCondition

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const Application = instance.Application;
  const sheet = Application.Sheets(1)
  const linkField = await sheet.FieldDescriptors(3)
  const linkGroups = await Application.AutoLinkGroups()
  const group = linkGroups.Add()
  const conditions = group.Conditions
  const fieldId =  sheet.FieldId("状态")
  conditions.Add(fieldId, [fieldId], "Field", "Equals")
  linkField.AutoLinkGroups = linkGroups
  linkField.IsAutoLink = true
  linkField.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main(){
  const sheet = Application.Sheets(1)
  const linkField = sheet.FieldDescriptors(3)
  const linkGroups = Application.AutoLinkGroups()
  const group = linkGroups.Add()
  const conditions = group.Conditions
  const fieldId =  sheet.FieldId("状态")
  const condition = conditions.Add(fieldId, [fieldId], "Field", "Equals")
  linkField.AutoLinkGroups = linkGroups
  linkField.IsAutoLink = true
  const result = linkField.Apply()
  console.log("####", result)
}
main()
```
 
# 343 API文档 / API / AutoLinkConditions / Delete

本页内容

# AutoLinkConditions.Delete(方法) ​

## 说明 ​

删除单向关联和双向关联字段自动匹配条件

## 语法 ​

表达式.Delete(Index)

表达式:AutoLinkConditions

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Index | 否 | Number | 索引 |

## 返回值 ​

无

## jsApi 示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const group = await app.Sheets(1).FieldDescriptors(2).AutoLinkGroups(1)
  const conditions = await groups.Conditions
  await conditions.Delete(1)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const group = Application.Sheets(1).FieldDescriptors(2).AutoLinkGroups(1)
  const conditions = groups.Conditions
  conditions.Delete(1)
 }
main()
```
 
# 344 API文档 / API / AutoLinkConditions / Item

本页内容

# AutoLinkConditions.Item(方法) ​

## 说明 ​

单向关联和双向关联字段自动匹配条件

## 语法 ​

表达式.Item(Index)

表达式:AutoLinkConditions

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Index | 否 | Number | 索引 |

## 返回值 ​

AutoLinkCondition

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const group = await app.Sheets(1).FieldDescriptors(2).AutoLinkGroups(1)
  const condition = await group.Conditions.Item(1)
  console.log(condition)
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const group = Application.Sheets(1).FieldDescriptors(2).AutoLinkGroups(1)
  const condition = group.Conditions.Item(1)
  console.log(condition)
 }
main()
```
 
# 345 API文档 / API / AutoLinkGroup / AutoLinkGroup对象

本页内容

# AutoLinkGroup (对象) ​

## 说明 ​

关联字段条件集合

## 方法 ​

## 属性 ​

- [Conditions](/documents/app-integration-dev/guide/dbsheet/Api/AutoLinkGroup_Conditions.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const group = await app.Sheets(1).FieldDescriptors(2).AutoLinkGroups(1)
  const Conditions = await group.Conditions
  console.log(await Conditions.Count)
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const group = Application.Sheets(1).FieldDescriptors(2).AutoLinkGroups(1)
  const Conditions = group.Conditions
  console.log(Conditions.Count)
 }
main()
```
 
# 346 API文档 / API / AutoLinkGroup / Conditions

本页内容

# AutoLinkGroup.Conditions(属性) ​

## 说明 ​

可读写

关联字段自动匹配条件集合

## 返回值 ​

AutoLinkConditions

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const group = await app.Sheets(1).FieldDescriptors(2).AutoLinkGroups(1)
  const Conditions = await group.Conditions
  console.log(await Conditions.Count)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const group = Application.Sheets(1).FieldDescriptors(2).AutoLinkGroups(1)
  const Conditions = groups.Conditions
  console.log(Conditions.Count)
}
main()
```
 
# 347 API文档 / API / AutoLinkGroups / AutoLinkGroups对象

本页内容

# AutoLinkGroups (对象) ​

## 说明 ​

关联字段的关系组集合，如果有多个关系组，则这些关系组是或的关系

## 方法 ​

- [Item](/documents/app-integration-dev/guide/dbsheet/Api/AutoLinkGroups_Item.html)
- [Add](/documents/app-integration-dev/guide/dbsheet/Api/AutoLinkGroups_Add.html)
- [Delete](/documents/app-integration-dev/guide/dbsheet/Api/AutoLinkGroups_Delete.html)

## 属性 ​

- [Count](/documents/app-integration-dev/guide/dbsheet/Api/AutoLinkGroups_Count.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const Application = instance.Application;
  const sheet = Application.Sheets(1)
  const linkField = await sheet.FieldDescriptors(3)
  const linkGroups = await Application.AutoLinkGroups()
  const group = linkGroups.Add()
  const conditions = group.Conditions
  const fieldId =  sheet.FieldId("状态")
  conditions.Add(fieldId, [fieldId], "Field", "Equals")
  linkField.AutoLinkGroups = linkGroups
  linkField.IsAutoLink = true
  linkField.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main(){
  const sheet = Application.Sheets(1)
  const linkField = sheet.FieldDescriptors(3)
  const linkGroups = Application.AutoLinkGroups()
  const group = linkGroups.Add()
  const conditions = group.Conditions
  const fieldId =  sheet.FieldId("状态")
  const condition = conditions.Add(fieldId, [fieldId], "Field", "Equals")
  linkField.AutoLinkGroups = linkGroups
  linkField.IsAutoLink = true
  const result = linkField.Apply()
  console.log("####", result)
}
main()
```
 
# 348 API文档 / API / AutoLinkGroups / Count

本页内容

# AutoLinkGroups.Count(属性) ​

## 说明 ​

可读 关联字段自动匹配条件组的数量

## 返回值 ​

Number

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const autoLinkGroups =  await app.Sheets(1).FieldDescriptors(2).AutoLinkGroups
 const count = await autoLinkGroups.Count
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const autoLinkGroups = Applicaiton.Sheets(1).FieldDescriptors(2).AutoLinkGroups
  const count = autoLinkGroups.Count
 }
main()
```
 
# 349 API文档 / API / AutoLinkGroups / Add

本页内容

# AutoLinkGroups.Add(方法) ​

## 说明 ​

增加关联字段匹配条件组

## 语法 ​

表达式.Add()

表达式:AutoLinkGroups

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |

## 返回值 ​

AutoLinkGroup

## jsApi 示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const autoLinkGroups =  await app.AutoLinkGroups()
 const count = await autoLinkGroups.Count
 if (count === 0) {
    const group = autoLinkGroups.Add()
 }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const autoLinkGroups = Application.AutoLinkGroups()
  const count = autoLinkGroups.Count
  if (count === 0) {
    const group = autoLinkGroups.Add()
  }
 }
main()```
```
 
# 350 API文档 / API / AutoLinkGroups / Delete

本页内容

# AutoLinkGroups.Delete(方法) ​

## 说明 ​

删除关系字段自动匹配条件组

## 语法 ​

表达式.Delete(Index)

表达式:AutoLinkGroups

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Index | 否 | [object Object] | 索引值 |

## 返回值 ​

## jsApi 示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const groups = await app.Sheets(1).FieldDescriptors(2).AutoLinkGroups
  groups.Delete(1)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const groups = Application.Sheets(1).FieldDescriptors(2).AutoLinkGroups
  groups.Delete(1)
}
main()```
```
 
# 351 API文档 / API / AutoLinkGroups / Item

本页内容

# AutoLinkGroups.Item(方法) ​

## 说明 ​

返回关联字段自动匹配条件组

## 语法 ​

表达式.Item(Index)

表达式:AutoLinkGroups

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Index | 否 | [object Object] | 索引值 |

## 返回值 ​

AutoLinkGroup

## jsApi 示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const autoLinkGroups =  await app.Sheets(1).FieldDescriptors(2).AutoLinkGroups
 const count = await autoLinkGroups.Count
 if (count > 0) {
    const group = await autoLinkGroups.Item(1)
 }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
 const autoLinkGroups =  Application.Sheets(1).FieldDescriptors(2).AutoLinkGroups
  const count = autoLinkGroups.Count
  if (count > 1) {
    const group = autoLinkGroups.Item(1)
  }
 }
main()```
```
 
# 352 API文档 / API / AutomationField / AutomationField对象

本页内容

# AutomationField (对象) ​

## 说明 ​

AutomationField 自动任务字段的属性，修改属性之后需要 FieldDescriptor 调用Apply()方法使得修改生效。如果非自动任务字段，无法设置相关属性

## 属性 ​

- [Type](/documents/app-integration-dev/guide/dbsheet/Api/AutomationField_Type.html)
- [TriggerField](/documents/app-integration-dev/guide/dbsheet/Api/AutomationField_TriggerField.html)
- [ContactField](/documents/app-integration-dev/guide/dbsheet/Api/AutomationField_ContactField.html)
- [ExecuteTime](/documents/app-integration-dev/guide/dbsheet/Api/AutomationField_ExecuteTime.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const Application = instance.Application;
  const sheet = Application.Sheets(1)
  const field = await sheet.FieldDescriptors(11)
  const automation = await field.Automation
  // 设置 自动化任务的属性
  automation.Type = "DueDateNotifyContact"
  automation.ExecuteTime = 3600
  // 更新字段
  field.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main(){
  const sheet = Application.Sheets(1)
  const field = sheet.FieldDescriptors(11)
  const automation = field.Automation
  // 设置 自动化任务的属性
  automation.Type = "DueDateNotifyContact"
  automation.ExecuteTime = 3600
  // 更新字段
  field.Apply()
}
main()
```
 
# 353 API文档 / API / AutomationField / ContactField

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 通知联系人 ​

AutomationField.ContactField(属性)

## 说明 ​

可读写

自动任务字段的通知联系人字段

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const Application = instance.Application;
  const sheet = Application.Sheets(1)
  const field = await sheet.FieldDescriptors(11)
  const automation = await field.Automation
  // 设置 自动化任务的属性
  automation.Type = "DueDateNotifyContact"
  automation.ContactField = await sheet.FieldId("联系人")
  // 更新字段
  field.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const sheet = Application.Sheets(1)
  const field = await sheet.FieldDescriptors(11)
  const automation = await field.Automation
  // 设置 自动化任务的属性
  automation.Type = "DueDateNotifyContact"
  automation.ContactField = sheet.FieldId("联系人")
  // 更新字段
  field.Apply()
}
main();
```
 
# 354 API文档 / API / AutomationField / ExecuteTime

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 触发时间 ​

AutomationField.ExecuteTime(属性)

## 说明 ​

可读写

当自动任务字段的类型为 DueDateNotifyContact，可设置指定触发时间通知联系人 格式为当前日期 + 执行时间, 时间为当天经过的秒数

## 返回值 ​

Number

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const Application = instance.Application;
  const sheet = Application.Sheets(1)
  const field = await sheet.FieldDescriptors(11)
  const automation = await field.Automation
  // 设置 自动化任务的属性
  automation.Type = "DueDateNotifyContact"
  automation.ExecuteTime = 3600
  // 更新字段
  field.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const sheet = Application.Sheets(1)
  const field = await sheet.FieldDescriptors(11)
  const automation = await field.Automation
  // 设置 自动化任务的属性
  automation.Type = "DueDateNotifyContact"
  automation.ExecuteTime = 3600
  // 更新字段
  field.Apply()
}
main();
```
 
# 355 API文档 / API / AutomationField / TriggerField

本页内容

# AutomationField.TriggerField(属性) ​

## 说明 ​

可读写

自动任务字段的触发字段，当Type为CheckedNotifyContact时，触发字段是复选框字段，当Type为DueDateNotifyContact时，触发字段为日期字段

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const Application = instance.Application;
  const sheet = Application.Sheets(1)
  const field = await sheet.FieldDescriptors("@自动任务")
  const automation = await field.Automation
  // 设置 自动化任务的属性
  automation.Type = "DueDateNotifyContact"
  automation.TriggerField = await sheet.FieldId("日期")
  // 更新字段
  field.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const sheet = Application.Sheets(1)
  const field = await sheet.FieldDescriptors(11)
  const automation = await field.Automation
  // 设置 自动化任务的属性
  automation.Type = "DueDateNotifyContact"
  automation.TriggerField = sheet.FieldId("日期")
  // 更新字段
  field.Apply()
}
main();
```
 
# 356 API文档 / API / AutomationField / Type

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 类型 ​

AutomationField.Type(属性)

## 说明 ​

可读写

自动任务字段的类型，有三种类型，参考枚举值 DbAutomationPresetType ，分别为CheckedNotifyContact，UpdatedNotifyContact，DueDateNotifyContact

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const Application = instance.Application;
    const contactField = await Application.FieldDescriptor("Contact","联系人字段")
    await Application.Sheets(1).FieldDescriptors.AddField(contactField)
    const dateField = await Application.FieldDescriptor("Date","日期字段")
    await Application.Sheets(1).FieldDescriptors.AddField(dateField)
    const newField = await Application.FieldDescriptor("Automations","自动任务字段")
    const automation = await newField.Automation
    automation.Type = "DueDateNotifyContact"
    automation.TriggerField = await dateField.Id
    automation.ContactField = await contactField.Id
    automation.ExecuteTime = 3600
    await Application.Sheets(1).FieldDescriptors.AddField(newField)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const contactField = FieldDescriptor("Contact","联系人字段")
    Sheets(1).FieldDescriptors.AddField(contactField)
    const dateField = FieldDescriptor("Date","日期字段")
    Sheets(1).FieldDescriptors.AddField(dateField)
    const newField = FieldDescriptor("Automations","自动任务字段")
    const automation = newField.Automation
    automation.Type = "DueDateNotifyContact"
    automation.TriggerField = dateField.Id
    automation.ContactField = contactField.Id
    automation.ExecuteTime = 3600
    Sheets(1).FieldDescriptors.AddField(newField)
}
main();
```
 
# 357 API文档 / API / ButtonField / ButtonField对象

本页内容

# ButtonField (对象) ​

## 说明 ​

ButtonField 按钮字段的属性，修改属性之后需要 FieldDescriptor 调用Apply()方法使得修改生效。如果非按钮字段，无法设置相关属性

## 属性 ​

- [BackgroundColor](/documents/app-integration-dev/guide/dbsheet/Api/ButtonField_BackgroundColor.html)
- [Icon](/documents/app-integration-dev/guide/dbsheet/Api/ButtonField_Icon.html)
- [Text](/documents/app-integration-dev/guide/dbsheet/Api/ButtonField_Text.html)
- [SuccessText](/documents/app-integration-dev/guide/dbsheet/Api/ButtonField_SuccessText.html)
- [TextColor](/documents/app-integration-dev/guide/dbsheet/Api/ButtonField_TextColor.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const Application = instance.Application
    const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@按钮")
    const prop = await fieldDescriptor.Button
    prop.Text = "Get it"
    prop.SuccessText = "SucccessText"
    prop.TextColor = "#ff00ff"
    prop.BackgroundColor = "#000000"
    prop.Icon = "calendar_check_in"
    fieldDescriptor.Apply()
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@按钮")
    const prop = field.Button
    prop.Text = "Get it"
    prop.SuccessText = "SucccessText"
    prop.TextColor = "#ff00ff"
    prop.BackgroundColor = "#000000"
    prop.Icon = "calendar_check_in"
    fieldDescriptor.Apply()
 }
main()
```
 
# 358 API文档 / API / ButtonField / BackgroundColor

本页内容

# ButtonField.BackgroundColor (属性) ​

## 说明 ​

可读写

按钮字段的填充颜色

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const Application = instance.Application
   const field = await Application.Sheets("数据表").FieldDescriptors("@按钮")
   const prop = await field.Button
   prop.BackgroundColor = "#1f00ff"
   field.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const field = Application.Sheets("数据表").FieldDescriptors("@按钮")
   const prop = field.Button
   prop.BackgroundColor = "#1f00ff"
   field.Apply()
}
main()
```
 
# 359 API文档 / API / ButtonField / Icon

本页内容

# FieldDescriptor.Icon(属性) ​

## 说明 ​

可读写

按钮字段的图标,可使用枚举值[DbButtonIcon](/documents/app-integration-dev/guide/dbsheet/Api/Enum_DbButtonIcon.html)

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const Application = instance.Application
   const field = await Application.Sheets("数据表").FieldDescriptors("@按钮")
   field.Button.Icon = "calendar_check_in"
   field.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const field = Application.Sheets("数据表").FieldDescriptors("@按钮")
   field.Button.Icon = Enum.DbButtonIcon.CalendarCheckIn
   field.Apply()
}
main()
```
 
# 360 API文档 / API / ButtonField / SuccessText

本页内容

# ButtonField.SuccessText(属性) ​

## 说明 ​

可读写

点击按钮字段后的提示文本

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const Application = instance.Application
   const field = await Application.Sheets("数据表").FieldDescriptors("@按钮")
   const prop = await field.Button
   prop.SuccessText = "点击按钮成功"
   field.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const field = Application.Sheets("数据表").FieldDescriptors("@按钮")
   const prop = field.Button
   prop.SuccessText = "点击按钮成功"
   field.Apply()
}
main()
```
 
# 361 API文档 / API / ButtonField / Text

本页内容

# ButtonField.Text(属性) ​

## 说明 ​

可读写

按钮字段 按钮显示的文本

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const Application = instance.Application
   const field = await Application.Sheets("数据表").FieldDescriptors("@按钮")
   const prop = await field.Button
   prop.Text = "Click"
   field.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const field = Application.Sheets("数据表").FieldDescriptors("@按钮")
   const prop = field.Button
   prop.Text = "Click"
   field.Apply()
}
main()
```
 
# 362 API文档 / API / ButtonField / TextColor

本页内容

# ButtonField.TextColor(属性) ​

## 说明 ​

可读写

按钮字段 按钮显示的文本颜色

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const Application = instance.Application
   const field = await Application.Sheets("数据表").FieldDescriptors("@按钮")
   const prop = await field.Button
   prop.TextColor = "#ff00ff"
   field.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const field = Application.Sheets("数据表").FieldDescriptors("@按钮")
   const prop = field.Button
   prop.TextColor = "#ff00ff"
   field.Apply()
}
main()
```
 
# 363 API文档 / API / CascadeField / CascadeField对象

本页内容

# CascadeField (对象) ​

## 说明 ​

CascadeField 级联字段的属性，修改属性之后需要 FieldDescriptor 调用Apply()方法使得修改生效。如果非级联字段，无法设置相关属性 级联字段最多设置为4级

## 属性 ​

- [IsDisplayAllLevel](/documents/app-integration-dev/guide/dbsheet/Api/CascadeField_IsDisplayAllLevel.html)
- [AllCascadeOption](/documents/app-integration-dev/guide/dbsheet/Api/CascadeField_AllCascadeOption.html)
- [Title](/documents/app-integration-dev/guide/dbsheet/Api/CascadeField_Title.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const Application = instance.Application
    const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@级联选项")
    const prop = await fieldDescriptor.Cascade
    prop.IsDisplayAllLevel = true
    fieldDescriptor.Apply()
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@级联选项")
    const prop = fieldDescriptor.Cascade
    prop.IsDisplayAllLevel = true
    fieldDescriptor.Apply()
 }
main()
```
 
# 364 API文档 / API / CascadeField / AllCascadeOption

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置选项 ​

CascadeField.AllCascadeOption(属性)

## 说明 ​

可读写

级联字段的选项，最多可以设置4级

## 返回值 ​

[CascadeOptions](/documents/app-integration-dev/guide/dbsheet/Api/CascadeOptions.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const Application = instance.Application;
  const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@级联选项")
  const prop = await fieldDescriptor.Cascade
  const allCascade =  await Application.CascadeOptions();
  const o1 = await allCascade.Add("级联1")
  o1.Children.Add("级联1_1")
  o1.Children.Add("级联1_2")
  const o2 = await allCascade.Add("级联2")
  o2.Children.Add("级联2_1")
  o2.Children.Add("级联2_2")
  const o3 = await allCascade.Add("级联3")
  o3.Children.Add("级联3_1")
  o3.Children.Add("级联3_2")
  const o4 = await allCascade.Add("级联4")
  o4.Children.Add("级联4_1")
  const o4_2 = o4.Children.Add("级联4_2")
  o4_2.Children.Add("级联4_2_1")
  o4_2.Children.Add("级联4_2_2")

  prop.AllCascadeOption = allCascade
  fieldDescriptor.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@级联选项")
  const prop = fieldDescriptor.Cascade
  const allCascade =  CascadeOptions();
  const o1 = allCascade.Add("级联1")
  o1.Children.Add("级联1_1")
  o1.Children.Add("级联1_2")
  const o2 = allCascade.Add("级联2")
  o2.Children.Add("级联2_1")
  o2.Children.Add("级联2_2")
  const o3 = allCascade.Add("级联3")
  o3.Children.Add("级联3_1")
  o3.Children.Add("级联3_2")
  const o4 = allCascade.Add("级联4")
  o4.Children.Add("级联4_1")
  const o4_2 = o4.Children.Add("级联4_2")
  o4_2.Children.Add("级联4_2_1")
  o4_2.Children.Add("级联4_2_2")
  
  prop.AllCascadeOption = allCascade
  fieldDescriptor.Apply()
}
main()
```
 
# 365 API文档 / API / CascadeField / IsDisplayAllLevel

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 显示完整的选择路径 ​

CascadeField.IsDisplayAllLevel(属性)

## 说明 ​

可读写

级联字段是否显示完整的选择路径

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const Application = instance.Application
   const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@级联选项")
   const prop = await fieldDescriptor.Cascade
   prop.IsDisplayAllLevel = true
   fieldDescriptor.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@级联选项")
   const prop = fieldDescriptor.Cascade
   prop.IsDisplayAllLevel = true
   fieldDescriptor.Apply()
}
main()
```
 
# 366 API文档 / API / CascadeField / Title

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 选项标题 ​

FieldDescriptor.Title(属性)

## 说明 ​

可读写

级联选项各级选项的标题，最多可以设置四级标题

## 返回值 ​

Array

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const Application = instance.Application;
  const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@级联选项")
  const prop = await fieldDescriptor.Cascade
  prop.Title = ["第一级标题","第二级标题","第三级标题"]
  fieldDescriptor.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const fieldDescriptor = Application.Sheets(1).FieldDescriptors(2)
  const prop = await fieldDescriptor.Cascade
  prop.Title = ["第一级标题","第二级标题","第三级标题"]
  fieldDescriptor.Apply()
}
main()
```
 
# 367 API文档 / API / CascadeOption / CascadeOption对象

本页内容

# CascadeOption (对象) ​

## 说明 ​

级联字段选项

## 方法 ​

## 属性 ​

- [Id](/documents/app-integration-dev/guide/dbsheet/Api/CascadeOption_Id.html)
- [Value](/documents/app-integration-dev/guide/dbsheet/Api/CascadeOption_Value.html)
- [Children](/documents/app-integration-dev/guide/dbsheet/Api/CascadeOption_Children.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const options =  await app.CascadeOptions();
 const o1 = await options.Add("test1")
 o1.Children.Add("child1")
 options.Add("test2")
 const desc =  await app.FieldDescriptor("Cascade","级联字段")
 desc.AllCascadeOption = options
 app.Sheets(1).FieldDescriptors.AddField(desc, 1)
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
 const options = CascadeOptions();
 options.Add("test1")
 options.Add("test2")
 const desc = FieldDescriptor("Cascade","级联字段")
 desc.AllCascadeOption = options
 Application.Sheets(1).FieldDescriptors.AddField(desc, 1)
 }
main()
```
 
# 368 API文档 / API / CascadeOption / Children

本页内容

# CascadeOption.Children(属性) ​

## 说明 ​

可读写

级联字段选项的子选项集合

## 返回值 ​

CascadeOptions

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const options =  await app.CascadeOptions();
 const o1 = await options.Add("test1")
 o1.Children.Add("child1")
 options.Add("test2")
 const desc =  await app.FieldDescriptor("Cascade","级联字段")
 desc.AllCascadeOption = options
 app.Sheets(1).FieldDescriptors.AddField(desc, 1)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
 const options = CascadeOptions();
 options.Add("test1")
 options.Add("test2")
 const desc = FieldDescriptor("Cascade","级联字段")
 desc.AllCascadeOption = options
 Application.Sheets(1).FieldDescriptors.AddField(desc, 1)
}
main()
```
 
# 369 API文档 / API / CascadeOption / Id

本页内容

# CascadeOption.Id(属性) ​

## 说明 ​

可读写

返回级联字段选项的Id

## 返回值 ​

string

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
   const app = instance.Application;
  const cascadeOptions = await app.Sheets(1).FieldDescriptors(2).AllCascadeOption
  console.log(await cascadeOptions.Item(1).Id)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const cascadeOptions = Application.Sheets(1).FieldDescriptors(2).AllCascadeOption
  console.log(cascadeOptions.Item(1).Id)
}
main()
```
 
# 370 API文档 / API / CascadeOption / Value

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# CascadeOption.Value(属性) ​

## 说明 ​

可读写

返回级联字段选项的值

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  await instance.ready();
   const app = instance.Application;
  const cascadeOptions = await app.Sheets(1).FieldDescriptors(2).AllCascadeOption
  console.log(await cascadeOptions.Item(1).Value)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const cascadeOptions = Application.Sheets(1).FieldDescriptors(2).AllCascadeOption
  console.log(cascadeOptions.Item(1).Value)
}
main()
```
 
# 371 API文档 / API / CascadeOptions / CascadeOptions对象

本页内容

# CascadeOptions (对象) ​

## 说明 ​

级联字段的选项集合

## 方法 ​

- [Item](/documents/app-integration-dev/guide/dbsheet/Api/CascadeOptions_Item.html)
- [Add](/documents/app-integration-dev/guide/dbsheet/Api/CascadeOptions_Add.html)

## 属性 ​

- [Count](/documents/app-integration-dev/guide/dbsheet/Api/CascadeOptions_Count.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const options =  await app.CascadeOptions();
 const o1 = await options.Add("test1")
 o1.Children.Add("child1")
 options.Add("test2")
 const desc =  await app.FieldDescriptor("Cascade","级联字段")
 desc.AllCascadeOption = options
 app.Sheets(1).FieldDescriptors.AddField(desc, 1)
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
 const options = CascadeOptions();
 options.Add("test1")
 options.Add("test2")
 const desc = FieldDescriptor("Cascade","级联字段")
 desc.AllCascadeOption = options
 Application.Sheets(1).FieldDescriptors.AddField(desc, 1)
 }
main()
```
 
# 372 API文档 / API / CascadeOptions / Count

本页内容

# CascadeOptions.Count(属性) ​

## 说明 ​

可读写

级联字段的选项数

## 返回值 ​

Number

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const descript = await app.Sheets(1).FieldDescriptors(2)
  const CascadeOption = await descript.AllCascadeOption
  const count = await CascadeOption.Count
  if (count === 0) {
    CascadeOption.Add("ADD 1")
    descript.Apply()
  }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const descript = Application.Sheets(1).FieldDescriptors(2)
  const CascadeOption = descript.AllCascadeOption
  const count = CascadeOption.Count
  if (count === 0) {
    CascadeOption.Add("ADD 1")
    descript.Apply()
  }
}
main()
```
 
# 373 API文档 / API / CascadeOptions / Add

本页内容

# CascadeOptions.Add(方法) ​

## 说明 ​

增加级联字段选项

## 语法 ​

表达式.Add(Value)

表达式:CascadeOptions

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Value | 否 | string | 选项值 |

## 返回值 ​

CascadeOption

## jsApi 示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const options = await app.CascadeOptions();
 const o1 = await options.Add("test1")
 o1.Children.Add("child1")
 options.Add("test2")
 const desc = await app.FieldDescriptor("Cascade","级联字段")
 desc.AllCascadeOption = options
 app.Sheets(1).FieldDescriptors.AddField(desc, 1)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
 const options = CascadeOptions();
 options.Add("test1")
 options.Add("test2")
 const desc = FieldDescriptor("Cascade","级联字段")
 desc.AllCascadeOption = options
 Application.Sheets(1).FieldDescriptors.AddField(desc, 1)
}
main()```
```
 
# 374 API文档 / API / CascadeOptions / Item

本页内容

# CascadeOptions.Item(方法) ​

## 说明 ​

返回级联字段的选项

## 语法 ​

表达式.Item(Index)

表达式:CascadeOptions

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Index | 否 | Number | 选项索引从1开始 |

## 返回值 ​

CascadeOption

## jsApi 示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
 const options = app.CascadeOptions();
 options.Add("test1")
 options.Add("test2")
 const desc = app.FieldDescriptor("Cascade","级联字段")
 desc.AllCascadeOption = options
 await Application.Sheets(1).FieldDescriptors.AddField(desc, 1)
 console.log(options.Item(1).Value)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
 const options = CascadeOptions();
 options.Add("test1")
 options.Add("test2")
 const desc = FieldDescriptor("Cascade","级联字段")
 desc.AllCascadeOption = options
 Application.Sheets(1).FieldDescriptors.AddField(desc, 1)
 console.log(options.Item(1).Value)
}
main()```
```
 
# 375 API文档 / API / ContactField / ContactField对象

本页内容

# ContactField (对象) ​

## 说明 ​

ContactField 联系人字段的属性，修改属性之后需要 FieldDescriptor 调用Apply()方法使得修改生效。如果非联系人字段，无法设置相关属性

## 属性 ​

- [IsSupportNotice](/documents/app-integration-dev/guide/dbsheet/Api/ContactField_IsSupportNotice.html)
- [IsSupportMulti](/documents/app-integration-dev/guide/dbsheet/Api/ContactField_IsSupportMulti.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const Application = instance.Application
    const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@联系人")
    const prop = await fieldDescriptor.Contact
    prop.IsSupportNotice = true
    prop.IsSupportMulti = false
    await fieldDescriptor.Apply()
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@联系人")
    const prop = fieldDescriptor.Contact
    prop.IsSupportNotice = true
    prop.IsSupportMulti = false
    fieldDescriptor.Apply()
 }
main()
```
 
# 376 API文档 / API / ContactField / IsSupportMulti

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 插入多个联系人 ​

ContactField.IsSupportMulti(属性)

## 说明 ​

可读写

字段类型为联系人时，通过此属性可以设置是否允许向单元格插入多个联系人

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const Application = instance.Application
    const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@联系人")
    const prop = await fieldDescriptor.Contact
    prop.IsSupportMulti = true
    fieldDescriptor.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@联系人")
    const prop = fieldDescriptor.Contact
    prop.IsSupportMulti = true
    fieldDescriptor.Apply()
}
main()
```
 
# 377 API文档 / API / ContactField / IsSupportNotice

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 发送通知 ​

ContactField.IsSupportNotice(属性)

## 说明 ​

可读写

字段类型为联系人时，通过此属性可以设置是否允许向新插入的联系人发送通知

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const Application = instance.Application
    const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@联系人")
    const prop = await fieldDescriptor.Contact
    prop.IsSupportNotice = true
    fieldDescriptor.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@联系人")
    const prop = fieldDescriptor.Contact
    prop.IsSupportNotice = true
    fieldDescriptor.Apply()
}
main()
```
 
# 378 API文档 / API / DBCellValue / DBCellValue对象

本页内容

# DBCellValue (对象) ​

## 说明 ​

单元格数据，读取单元格值时， 如果是复杂的字段数据，则会返回DBCellValue, 比如：多选项/超链接/联系人/级联字段/关联字段/地址字段/图片与附件。 当通过RecordRange.Value 设置复杂的字段数据时，需要构造DBCellValue作为参数设置到单元格

```javascript
ActiveView.RecordRange([5,6], "@地址").Value = DBCellValue({districts:["广东省","珠海市","香洲区"],detail:"云海路"})
```

## 方法 ​

## 属性 ​

- [Value](/documents/app-integration-dev/guide/dbsheet/Api/DBCellValue_Value.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const dbCellValue = await app.Sheets(1).Views(1).RecordRange(1,"@图片和附件").Value
 const attments = await dbCellValue.Value
 console.log(await attments[0].FileName)
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
 const dbCellValue = Application.Sheets(1).Views(1).RecordRange(1,"@图片和附件").Value
 const attments = dbCellValue.Value
 console.log(attments[0].FileName)
 }
main()
```
 
# 379 API文档 / API / DBCellValue / Value

本页内容

# DBCellValue.Value(属性) ​

## 说明 ​

返回单元格的值 地址字段类型: 通过DBCellValue() 生成字段的数据

```javascript
DBCellValue({districts:["广东省","珠海市","香洲区"],detail:"前岛环路xxxx号"})
```

级联字段类型：通过DBCellValue() 生成字段的数据

```javascript
DBCellValue({districts:["广东省","珠海市","香洲区"]})
```

超链接字段类型：

```javascript
DBCellValue({address:"wps.cn", display:"wps"})
```

关联字段类型: 参数传入关联的记录id

```javascript
DBCellValue(["b","V"])
```

多选项类型：

```javascript
DBCellValue(["未,开始","进行中"])
```

图片与附件字段：可以传入包含 URL/base64编码的图片/云文档 的数组，支持多个附件。 注意：由于脚本有运行时长限制，附件较大/或者较多时会导致超时，设置失败

```javascript
DBCellValue([{fileData: url/base64, fileName: ""}])
```
 
# 380 API文档 / API / DateField / DateField对象

本页内容

# DateField (对象) ​

## 说明 ​

DateField 日期字段的属性，修改属性之后需要 FieldDescriptor 调用Apply()方法使得修改生效。如果非日期字段，无法设置相关属性

## 属性 ​

- [IsShowWeek](/documents/app-integration-dev/guide/dbsheet/Api/DateField_IsShowWeek.html)
- [IsShowTime](/documents/app-integration-dev/guide/dbsheet/Api/DateField_IsShowTime.html)
- [IsShowHoliday](/documents/app-integration-dev/guide/dbsheet/Api/DateField_IsShowHoliday.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const Application = instance.Application
    const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@日期")
    const prop = await fieldDescriptor.Date
    prop.IsShowWeek = true
    prop.IsShowTime = true
    prop.IsShowHoliday = false
    fieldDescriptor.Apply()
    console.log(await prop.IsShowWeek)
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@日期")
    const prop = fieldDescriptor.Date
    prop.IsShowWeek = true
    prop.IsShowTime = true
    prop.IsShowHoliday = false
    fieldDescriptor.Apply()
 }
main()
```
 
# 381 API文档 / API / DateField / IsShowHoliday

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 面板标记休息日 ​

DateField.IsShowHoliday(属性)

## 说明 ​

可读写

只对日期字段有效，是否在选择日期的面板上标识节假日

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const Application = instance.Application
  const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@日期")
  const prop = await fieldDescriptor.Date
  prop.IsShowHoliday = true
  fieldDescriptor.Apply()
  console.log(await prop.IsShowHoliday)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@日期")
  const prop = fieldDescriptor.Date
  prop.IsShowHoliday = true
  fieldDescriptor.Apply()
}
main()
```
 
# 382 API文档 / API / DateField / IsShowTime

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 显示时间 ​

DateField.IsShowTime(属性)

## 说明 ​

可读写

字段类型为 日期 时，通过此属性可以设置是否显示时间

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const Application = instance.Application
  const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@日期")
  const prop = await fieldDescriptor.Date
  prop.IsShowTime = true
  fieldDescriptor.Apply()
  console.log(await prop.IsShowTime)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@日期")
  const prop = fieldDescriptor.Date
  prop.IsShowTime = true
  fieldDescriptor.Apply()
}
main()
```
 
# 383 API文档 / API / DateField / IsShowWeek

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 是否显示星期 ​

DateField.IsShowWeek(属性)

## 说明 ​

可读写

字段类型为 日期 时，通过此属性设置是否显示星期

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const Application = instance.Application
  const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@日期")
  const prop = await fieldDescriptor.Date
  prop.IsShowWeek = true
  fieldDescriptor.Apply()
  console.log(await prop.IsShowWeek)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@日期")
  const prop = fieldDescriptor.Date
  prop.IsShowWeek = true
  fieldDescriptor.Apply()
}
main()
```
 
# 384 API文档 / API / Field / Field对象

本页内容

- 说明
- 方法
- 属性
- 浏览器环境示例
- 脚本编辑器示例

# Field (对象) ​

## 说明 ​

网格视图的列属性，提供修改列宽和隐藏/显示列的方法

## 方法 ​

- [Copy](/documents/app-integration-dev/guide/dbsheet/Api/Field_Copy.html)
- [Move](/documents/app-integration-dev/guide/dbsheet/Api/Field_Move.html)

## 属性 ​

- [Id](/documents/app-integration-dev/guide/dbsheet/Api/Field_Id.html)
- [Visible](/documents/app-integration-dev/guide/dbsheet/Api/Field_Visible.html)
- [Name](/documents/app-integration-dev/guide/dbsheet/Api/Field_Name.html)
- [Type](/documents/app-integration-dev/guide/dbsheet/Api/Field_Type.html)
- [Width](/documents/app-integration-dev/guide/dbsheet/Api/Field_Width.html)
- [FieldDescriptor](/documents/app-integration-dev/guide/dbsheet/Api/Field_FieldDescriptor.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const fields = await app.Sheets(1).Views(1).Fields
 console.log(await fields.Item(1).Id)
 }
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 const fields = Application.Sheets(1).Views(1).Fields
 console.log(fields.Item(1).Id)
}
main()
```
 
# 385 API文档 / API / Field / FieldDescriptor

本页内容

# Field.FieldDescriptor(属性) ​

## 说明 ​

只读 字段的数据定义

## 返回值 ​

[FieldDescriptor](/documents/app-integration-dev/guide/dbsheet/Api/FieldDescriptor.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const field = await app.Sheets(1).Views(1).Fields(1);
    const desc = await field.FieldDescriptor;
    console.log(await desc.Name);
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
    const field = Application.Sheets(1).Views(1).Fields(1);
    const desc = field.FieldDescriptor;
    console.log(desc.Name);
}
main();
```
 
# 386 API文档 / API / Field / Id

本页内容

# Field.Id(属性) ​

## 说明 ​

只读 视图字段的Id

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const fields = await app.Sheets(1).Views(1).Fields
 console.log(await fields.Item(1).Id)
 }
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 const fields = Application.Sheets(1).Views(1).Fields
 console.log(fields.Item(1).Id)
}
main()
```
 
# 387 API文档 / API / Field / Name

本页内容

# Field.Name(属性) ​

## 说明 ​

只读 视图字段的Name

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const field = await app.Sheets(1).Views(1).Fields(1)
 const name = await field.Name
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 const field = Application.Sheets(1).Views(1).Fields(1)
 const name = field.Name
 console.log(name)
}
main()
```
 
# 388 API文档 / API / Field / Type

本页内容

# Field.Type(属性) ​

## 说明 ​

只读 视图字段的Type

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const field = await app.Sheets(1).Views(1).Fields(1)
 const type = await field.Type
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 const field = Application.Sheets(1).Views(1).Fields(1)
 const type = field.Type
 console.log(type)
}
main()
```
 
# 389 API文档 / API / Field / Visible

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器示例

# 隐藏/显示字段 ​

Field.Visible(属性)

## 说明 ​

可读写

视图的字段是否可见

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const field = await app.Sheets(1).Views(1).Fields(1)
 field.Visible = false
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 const field = Application.Sheets(1).Views(1).Fields(1)
 const visible = field.Visible
 console.log(visible)
}
main()
```
 
# 390 API文档 / API / Field / Width

本页内容

# Field.Width(属性) ​

## 说明 ​

只读 视图字段的Width

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const field = await app.Sheets(1).Views(1).Fields(1)
 const Width = await field.Width
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 const field = Application.Sheets(1).Views(1).Fields(1)
 const width = field.Width
 console.log(width)
}
main()
```
 
# 391 API文档 / API / Field / Copy

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器示例

# 复制字段 ​

## 说明 ​

复制当前字段到指定位置

## 语法 ​

表达式.Copy(Before,After)

表达式:Field

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Before | 否 | String/Number | 在Before之前插入复制字段 |
| After | 否 | String/Number | 在After之后插入复制字段 |

## 返回值 ​

Self

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const field = await app.Sheets(1).Views(1).Fields(1)
 field.Copy(3)
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 const field = Application.Sheets(1).Views(1).Fields(1)
 field.Copy(3)
}
main()
```
 
# 392 API文档 / API / Field / Move

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器示例

# 移动字段 ​

## 说明 ​

移动字段到指定位置

## 语法 ​

表达式.Move(Before,After)

表达式:Field

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Before | 否 | String/Number | 移动到Before字段之前 |
| After | 否 | String/Number | 移动到After字段之后 |

## 返回值 ​

ApiResult

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const field = await app.Sheets(1).Views(1).Fields(1)
 field.Move(3)
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 const field = Application.Sheets(1).Views(1).Fields(1)
 field.Move(3)
}
main()
```
 
# 393 API文档 / API / FieldDescriptor / FieldDescriptor对象

本页内容

# FieldDescriptor (对象) ​

## 说明 ​

FieldDescriptor 描述了字段的属性，可以通过SetType修改字段的类型，修改属性之后需要主动调用Apply()方法使得修改生效。 每次读取FieldDescriptor都会重新生成数据，所以修改前需要记录下当前的FieldDescriptor，才能正确调用Apply()

```javascript
const fieldDescriptor = Sheets(1).FieldDescriptors(1)
// 设置属性
// ....
fieldDescriptor.Apply()
```

不同的字段类型有不同的属性设置，不是对应的字段类型，获取相关的属性时会返回null，无法正常设置

| 字段类型 | 特有的字段属性 |
| --- | --- |
| 按钮字段 | [Button](/documents/app-integration-dev/guide/dbsheet/Api/ButtonField.html) |
| 地址字段 | [Address](/documents/app-integration-dev/guide/dbsheet/Api/AddressField.html) |
| 级联字段 | [Cascade](/documents/app-integration-dev/guide/dbsheet/Api/CascadeField.html) |
| 联系人字段 | [Contact](/documents/app-integration-dev/guide/dbsheet/Api/ContactField.html) |
| 日期字段 | [Date](/documents/app-integration-dev/guide/dbsheet/Api/DateField.html) |
| 最后修改人/最后修改时间 | [Watch](/documents/app-integration-dev/guide/dbsheet/Api/WatchedField.html) |
| 公式字段 | [Formula](/documents/app-integration-dev/guide/dbsheet/Api/FormulaField.html) |
| 引用/查找引用/统计 | [Lookup](/documents/app-integration-dev/guide/dbsheet/Api/LookupField.html) |
| 单向关联/双向关联 | [Link](/documents/app-integration-dev/guide/dbsheet/Api/LinkField.html) |
| 自动任务 | [Automation](/documents/app-integration-dev/guide/dbsheet/Api/AutomationField.html) |
| 图片和附件 | [Attachment](/documents/app-integration-dev/guide/dbsheet/Api/AttachmentField.html) |
| 超链接 | [Url](/documents/app-integration-dev/guide/dbsheet/Api/UrlField.html) |
| 数字 | [Number](/documents/app-integration-dev/guide/dbsheet/Api/NumberField.html) |
| 单选项/多选项 | [Select](/documents/app-integration-dev/guide/dbsheet/Api/SelectField.html) |
| 等级 | [Rating](/documents/app-integration-dev/guide/dbsheet/Api/RatingField.html) |

## 方法 ​

- [Apply](/documents/app-integration-dev/guide/dbsheet/Api/FieldDescriptor_Apply.html)
- [Delete](/documents/app-integration-dev/guide/dbsheet/Api/FieldDescriptor_Delete.html)

## 属性 ​

- [Id](/documents/app-integration-dev/guide/dbsheet/Api/FieldDescriptor_Id.html)
- [Type](/documents/app-integration-dev/guide/dbsheet/Api/FieldDescriptor_Type.html)
- [Name](/documents/app-integration-dev/guide/dbsheet/Api/FieldDescriptor_Name.html)
- [Description](/documents/app-integration-dev/guide/dbsheet/Api/FieldDescriptor_Description.html)
- [IsSyncField](/documents/app-integration-dev/guide/dbsheet/Api/FieldDescriptor_IsSyncField.html)
- [DefaultVal](/documents/app-integration-dev/guide/dbsheet/Api/FieldDescriptor_DefaultVal.html)
- [DefaultValType](/documents/app-integration-dev/guide/dbsheet/Api/FieldDescriptor_DefaultValType.html)
- [NumberFormat](/documents/app-integration-dev/guide/dbsheet/Api/FieldDescriptor_NumberFormat.html)
- [IsValueUnique](/documents/app-integration-dev/guide/dbsheet/Api/FieldDescriptor_IsValueUnique.html)
- [Button](/documents/app-integration-dev/guide/dbsheet/Api/ButtonField.html)
- [Address](/documents/app-integration-dev/guide/dbsheet/Api/AddressField.html)
- [Cascade](/documents/app-integration-dev/guide/dbsheet/Api/CascadeField.html)
- [Contact](/documents/app-integration-dev/guide/dbsheet/Api/ContactField.html)
- [Date](/documents/app-integration-dev/guide/dbsheet/Api/DateField.html)
- [Watch](/documents/app-integration-dev/guide/dbsheet/Api/WatchedField.html)
- [Formula](/documents/app-integration-dev/guide/dbsheet/Api/FormulaField.html)
- [Lookup](/documents/app-integration-dev/guide/dbsheet/Api/LookupField.html)
- [Link](/documents/app-integration-dev/guide/dbsheet/Api/LinkField.html)
- [Automation](/documents/app-integration-dev/guide/dbsheet/Api/AutomationField.html)
- [Attachment](/documents/app-integration-dev/guide/dbsheet/Api/AttachmentField.html)
- [Url](/documents/app-integration-dev/guide/dbsheet/Api/UrlField.html)
- [Number](/documents/app-integration-dev/guide/dbsheet/Api/NumberField.html)
- [Select](/documents/app-integration-dev/guide/dbsheet/Api/SelectField.html)
- [Rating](/documents/app-integration-dev/guide/dbsheet/Api/RatingField.html)

## 事件 ​

- [OnUpdate](/documents/app-integration-dev/guide/dbsheet/Api/FieldDescriptor_OnUpdate.html)
- [OnDelete](/documents/app-integration-dev/guide/dbsheet/Api/FieldDescriptor_OnDelete.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const fieldDescriptor = await app.Sheets(1).FieldDescriptors(1)
    fieldDescriptor.Name = "修改字段名"
    fieldDescriptor.Apply()
    console.log(await fieldDescriptor.Name)
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const fieldDescriptor = Application.Sheets(1).FieldDescriptors(1)
    fieldDescriptor.Name = "修改字段名"
    fieldDescriptor.Apply()
    console.log(fieldDescriptor.Name)
 }
main()
```
 
# 394 API文档 / API / FieldDescriptor / DefaultVal

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置默认值 ​

FieldDescriptor.DefaultVal(属性)

## 说明 ​

可读写

设置默认值

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
   const app = instance.Application;
   const field = await WPSOpenApi.Application.Sheets(1).FieldDescriptors(2)
   field.DefaultVal = "1"
   field.Apply()
   console.log(await field.DefaultVal)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const field = Application.Sheets(1).FieldDescriptors(2)
   field.DefaultVal = "1"
   field.Apply()
   console.log(field.DefaultVal)
}
main()
```
 
# 395 API文档 / API / FieldDescriptor / DefaultValType

本页内容

# FieldDescriptor.DefaultValType(属性) ​

## 说明 ​

可读写

设置默认值类型

## 返回值 ​

Enum.DbFieldDefaultValueType

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
   const app = instance.Application;
   const field = await WPSOpenApi.Application.Sheets(1).FieldDescriptors(2)
   field.DefaultValType = app.Enum.DbFieldDefaultValueType.Normal
   field.Apply()
   console.log(await field.DefaultValType)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const field = Application.Sheets(1).FieldDescriptors(2)
   field.DefaultValType = Enum.DbFieldDefaultValueType.Normal
   field.Apply()
   console.log(field.DefaultValType)
}
main()
```
 
# 396 API文档 / API / FieldDescriptor / Description

本页内容

# FieldDescriptor.Description(属性) ​

## 说明 ​

可读写

字段的描述信息

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
   const app = instance.Application;
   const description = await app.Sheets(1).FieldDescriptors(2).Description;
   console.log(description)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const description = Application.Sheets(1).FieldDescriptors(2).Description
  console.log(description)
}
main()
```
 
# 397 API文档 / API / FieldDescriptor / Id

本页内容

# FieldDescriptor.Id(属性) ​

## 说明 ​

只读 返回字段的ID，用于标识字段

## 返回值 ​

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
   const field = await WPSOpenApi.Application.Sheets(1).FieldDescriptors(2)
   console.log(await field.Id)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const field = Application.Sheets(1).FieldDescriptors(2)
   console.log(field.Id)
}
main()
```
 
# 398 API文档 / API / FieldDescriptor / IsSyncField

本页内容

# FieldDescriptor.SyncField(属性) ​

## 说明 ​

只读 标记字段是否同步字段

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const field = await app.Sheets(1).FieldDescriptors(2)
  console.log(await field.SyncField)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const field = Application.Sheets(1).FieldDescriptors(2)
  console.log(field.SyncField) 
}
main()
```
 
# 399 API文档 / API / FieldDescriptor / IsValueUnique

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 禁止录入重复值 ​

FieldDescriptor.ValueUnique(属性)

## 说明 ​

可读写

是否唯一值，禁止重复录入

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const field = await app.Sheets(1).FieldDescriptors(2)
  console.log(await field.ValueUnique)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const field = Application.Sheets(1).FieldDescriptors(2)
  console.log(field.ValueUnique)
}
main()
```
 
# 400 API文档 / API / FieldDescriptor / Name

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置字段名称 ​

FieldDescriptor.Name(属性)

## 说明 ​

可读写

字段名称

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const field = await app.Sheets(1).FieldDescriptors(2)
  field.Name = "字段名"
  field.Apply()

  console.log(await field.Name)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const field = Application.Sheets(1).FieldDescriptors(2)
  field.Name = "字段名"
  field.Apply()

  console.log(await field.Name)
}
main()
```
 
# 401 API文档 / API / FieldDescriptor / NumberFormat

本页内容

- 说明
    - 货币
    - 数值
    - 百分比
    - 日期时间
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# FieldDescriptor.NumberFormat(属性) ​

## 说明 ​

可读写

数字格式，可以设置日期/时间/数值等字段的显示数字格式 可以设置内置的格式

| 字段类型 | 内建格式 |
| --- | --- |
| 货币 | '￥','$','"€"','£' |

### 货币 ​

| 显示结果 | 内建格式 |
| --- | --- |
| '￥' | '￥' |
| '$' | '$' |
| '€' | '\"€\"' |
| '£' | '£' |

### 数值 ​

| 显示结果 | 内建格式 |
| --- | --- |
| '1234' | '0 /  ' |
| '1234.1' | '0.0 /  ' |
| '1234.10 ' | '0.00 /  ' |
| '1234.100' | '0.000 /  ' |
| '1234.1000' | '0.0000 /  ' |
| '1,234' | '#,##0 /  ' |
| '1,234.1' | '#,##0.0 /  ' |
| '1,234.10' | '#,##0.00 /  ' |
| '1,234.100' | '#,##0.000 /  ' |
| '1,234.1000' | '#,##0.0000 /  ' |

### 百分比 ​

| 显示结果 | 内建格式 |
| --- | --- |
| '12%' | '0%' |
| '12.0%' | '0.0%' |
| '12.00% ' | '0.00%' |
| '12.000%' | '0.000%' |
| '12.0000%' | '0.0000%' |

### 日期时间 ​

| 显示结果 | 内建格式 |
| --- | --- |
| 2023/04/18 10：09 | 'yyyy/mm/dd h:mm' |

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const field = await app.Sheets(1).FieldDescriptors(2);
    await field.NumberFormat;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const field = Application.Sheets(1).FieldDescriptors(2);
    field.NumberFormat = '0.0000%';
    field.Apply();
}
main();
```
 
# 402 API文档 / API / FieldDescriptor / Type

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置字段类型 ​

FieldDescriptor.Type

## 说明 ​

可读写

返回当前字段类型

## 返回值 ​

| 字段类型 | 描述 |
| --- | --- |
| ID | 身份证 |
| Phone | 电话 |
| Email | 电子邮箱 |
| Url | 超链接 |
| Checkbox | 复选框 |
| SingleSelect | 单选项 |
| MultipleSelect | 多选项 |
| Rating | 等级 |
| Complete | 进度条 |
| CellPicture | 单元格图片 |
| Contact | 联系人 |
| Attachment | 附件 |
| Note | 富文本字段，备注 |
| Link | 关联 |
| OneWayLink | 单向关联 |
| Lookup | 引用 |
| Address | 地址,特殊级联字段 |
| Cascade | 级联 |
| Automations | 触发器 |
| AutoNumber | 编号 |
| CreatedBy | 创建者 |
| CreatedTime | 创建时间 |
| LastModifiedBy | 最后修改者 |
| LastModifiedTime | 最后修改时间 |
| Formula | 公式 |

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
   const app = instance.Application;
   const field = await app.Sheets(1).FieldDescriptors(2)
   console.log(await field.Type)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const field = Application.Sheets(1).FieldDescriptors(2)
   console.log(field.Type)
}
main()
```
 
# 403 API文档 / API / FieldDescriptor / Apply

本页内容

# FieldDescriptor.Apply(方法) ​

## 说明 ​

修改属性后调用Apply()方法使修改生效

## 语法 ​

表达式.Apply()

表达式:FieldDescriptor

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |

## 返回值 ​

ApiResult

## jsApi 示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const fieldId = app.Sheets(1).FieldId("修改前字段名")
    const fieldDescriptor = app.Sheets(1).FieldDescriptors(fieldId)
    fieldDescriptor.Name = "修改字段名"
    fieldDescriptor.Apply()
    console.log(fieldDescriptor.Name)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const app = Application;
    const fieldId = app.Sheets(1).FieldId("修改前字段名")
    const fieldDescriptor = app.Sheets(1).FieldDescriptors(fieldId)
    fieldDescriptor.Name = "修改字段名"
    fieldDescriptor.Apply()
    console.log(fieldDescriptor.Name)
}
main()
```
 
# 404 API文档 / API / FieldDescriptor / Delete

本页内容

- 说明
- 语法
- 参数
- 返回值
- jsApi 示例
- 脚本编辑器 示例

# 删除字段 ​

## 说明 ​

删除字段

## 语法 ​

表达式.Delete(RemoveReversedLink)

表达式:FieldDescriptor

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| RemoveReversedLink | 否 | Boolean |  |

## 返回值 ​

Boolean

## jsApi 示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
   const app = instance.Application;
   const field = await WPSOpenApi.Application.Sheets(1).FieldDescriptors(2)
   field.Delete()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const app = Application;
   const field = Application.Sheets(1).FieldDescriptors(2)
   field.Delete()
}
main()
```
 
# 405 API文档 / API / FieldDescriptor / OnDelete

本页内容

- 说明
- 语法
- 参数
- 返回值
- 事件返回数据
- 事件返回数据示例
- 浏览器环境示例
- 脚本编辑器 示例

# 监听删除字段的事件 ​

FieldDescriptor.OnDelete(方法)

## 说明 ​

为 FieldDescriptor 添加 Delete 事件,当删除 FieldDescriptor 时触发。注意在脚本编辑器中使用时，脚本运行结束就会退出运行，这时可能回调无法被正常触发

## 语法 ​

表达式.OnDelete(Callback)

表达式: FieldDescriptor

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Callback | 是 | func | 指定事件发生时的回调方法 ; const eventContext = await FieldDescriptor.OnDelete(()=&gt;{ ... }) |

## 返回值 ​

EventContext

## 事件返回数据 ​

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| sheetId | Number | 表的 Id |
| fieldId | String | 字段的 Id |
| fieldIds | Array | 字段集合的 Ids |

## 事件返回数据示例 ​

```javascript
{
    fieldId: "C"
    fieldIds: ['C']
    sheetId: 1
}
```

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const field = await app.Sheets(1).FieldDescriptors(2);
    let eventContext;
    eventContext = await field.OnDelete(data => {
        console.log(data);
        // 取消事件监听
        eventContext.Destroy();
    });
    await field.Delete();
    //这里会执行OnDelete的回调
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const field = Application.Sheets(1).FieldDescriptors(2);
    let eventContext;
    eventContext = field.OnDelete(data => {
        console.log(data);
        // 取消事件监听
        eventContext.Destroy();
    });
    field.Delete();
    field.Delete();
    //这里会执行OnDelete的回调
}
main();
```
 
# 406 API文档 / API / FieldDescriptor / OnUpdate

本页内容

# 监听修改字段的事件 ​

FieldDescriptor.OnUpdate(方法)

## 说明 ​

为 FieldDescriptor 添加 Update 事件,当更新 FieldDescriptor 时触发。注意在脚本编辑器中使用时，脚本运行结束就会退出运行，这时可能回调无法被正常触发

## 语法 ​

表达式.OnUpdate(Callback)

表达式: FieldDescriptor

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Callback | 是 | func | 指定事件发生时的回调方法 ; const eventContext = await FieldDescriptor.OnUpdate(()=&gt;{ ... }) |

## 返回值 ​

EventContext

## 事件返回数据 ​

FieldDescriptor

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const fieldId = app.Sheets(1).FieldId('修改前字段名');
    const fieldDescriptor = app.Sheets(1).FieldDescriptors(fieldId);
    let eventContext
    eventContext = await fieldDescriptor.OnUpdate(data => {
        console.log(data);
        // 取消事件监听
        eventContext.Destroy()
    });

    fieldDescriptor.Name = '修改字段名';
    fieldDescriptor.Apply();
    //这里会执行OnUpdate的回调
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const fieldDescriptor = Application.Sheets(1).FieldDescriptors(1);
    let eventContext
    eventContext = fieldDescriptor.OnUpdate(data => {
        console.log(data);
        // 取消事件监听
        eventContext.Destroy()
    });
    fieldDescriptor.Name = '修改字段名';
    fieldDescriptor.Apply();
    //这里会执行OnUpdate的回调
}
main();
```
 
# 407 API文档 / API / FieldDescriptors / FieldDescriptors对象

本页内容

- 说明
- 方法
- 属性
- 事件
- 浏览器环境示例
- 脚本编辑器 示例

# FieldDescriptors (对象) ​

## 说明 ​

字段描述的集合，保存了文档所有的字段的信息

## 方法 ​

- [Item](/documents/app-integration-dev/guide/dbsheet/Api/FieldDescriptors_Item.html)
- [FieldDescriptor](/documents/app-integration-dev/guide/dbsheet/Api/FieldDescriptors_FieldDescriptor.html)
- [AddField](/documents/app-integration-dev/guide/dbsheet/Api/FieldDescriptors_AddField.html)
- [Delete](/documents/app-integration-dev/guide/dbsheet/Api/FieldDescriptors_Delete.html)

## 属性 ​

- [Count](/documents/app-integration-dev/guide/dbsheet/Api/FieldDescriptors_Count.html)

## 事件 ​

- [OnCreate](/documents/app-integration-dev/guide/dbsheet/Api/FieldDescriptors_OnCreate.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const fieldDescriptors = await app.Sheets(1).FieldDescriptors
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main(){
const sheet = Application.Sheets(1)
const sheet2 = Application.Sheets(2)
const lookupFieldId = sheet2.FieldId("文本")
// 创建关联字段到sheet2
const linkField = Application.FieldDescriptor(Enum.DbSheetFieldType.OneWayLink, "关联字段到sheet2")
linkField.LinkSheet = sheet2.StId
linkField.IsAutoLink = false
/*
const linkGroups = Application.AutoLinkGroups()
const group = linkGroups.Add()
const conditions = group.Conditions
conditions.Add(lookupFieldId, [lookupFieldId], Enum.DbAutolinkCondType.Field, Enum.DbAutolinkCondType.DbFilterCriteriaOpType.Equals, )
linkField.AutoLinkGroups = linkGroups
*/
sheet.FieldDescriptors.AddField(linkField)

// 创建引用字段
const descriptor = Application.FieldDescriptor(Enum.DbSheetFieldType.Lookup, "引用关联字段到sheet2")
descriptor.LinkFieldId = sheet.FieldId("关联")
descriptor.LookupFieldId = lookupFieldId
sheet.FieldDescriptors.AddField(descriptor, 6)
}
main()
```
 
# 408 API文档 / API / FieldDescriptors / Count

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# FieldDescriptors.Count(属性) ​

## 说明 ​

可读写

字段描述的个数

## 返回值 ​

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
   const app = instance.Application;
   const count = await app.Sheets(1).FieldDescriptors.Count
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const count = Application.Sheets(1).FieldDescriptors.Count
}
main()
```
 
# 409 API文档 / API / FieldDescriptors / AddField

本页内容

- 说明
- 语法
- 参数
- 返回值
- jsApi 示例
- 脚本编辑器 示例

# 新增字段 ​

## 说明 ​

往表中新增新的字段

## 语法 ​

表达式.AddField(FieldDescriptor, Index)

表达式:FieldDescriptors

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| FieldDescriptor | 是 | FieldDescriptor | 字段属性 |
| Index | 否 | string/number | Index为string时表示字段ID，number时表示字段索引，插入位置，未指定时插入到末尾 |

## 返回值 ​

ApiResult

## jsApi 示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const desc = await app.FieldDescriptor("Rating","等级字段")
  desc.MaxRating = 2
  await app.Sheets(1).FieldDescriptors.AddField(desc, 1)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const desc = FieldDescriptor("Rating","等级字段")
  desc.MaxRating = 2
  await Application.Sheets(1).FieldDescriptors.AddField(desc, 1)
}
main()
```
 
# 410 API文档 / API / FieldDescriptors / Delete

本页内容

- 说明
- 语法
- 参数
- 返回值
- jsApi 示例
- 脚本编辑器 示例

# FieldDescriptors.Delete(方法) ​

## 说明 ​

删除字段

## 语法 ​

表达式.Delete(FieldId, RemoveReversedLink)

表达式:FieldDescriptors

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| FieldId | 是 | string | 字段Id,可以通过字段名获取Sheet.FieldId(fieldName) |
| RemoveReversedLink | 否 | boolean | 是否同时删除引用 |

## 返回值 ​

Boolean

## jsApi 示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const fieldId = await app.Sheets(1).FieldId("数字")
  app.Sheets(1).FieldDescriptors.Delete(fieldId)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const fieldId = Application.Sheets(1).FieldId("数字")
  app.Sheets(1).FieldDescriptors.Delete(fieldId)
}
main()
```
 
# 411 API文档 / API / FieldDescriptors / FieldDescriptor

本页内容

- 说明
- 语法
- 参数
- 返回值
- jsApi 示例
- 脚本编辑器 示例

# FieldDescriptors.FieldDescriptor(方法) ​

## 说明 ​

FieldDescriptor对象

## 语法 ​

表达式.FieldDescriptor(Type,Name)

表达式:FieldDescriptors

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Type | 是 |  | FieldDescriptor的类型 |
| Name | 是 | string | FieldDescriptor的名称 |

## 返回值 ​

FieldDescriptor

## jsApi 示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const fieldDescs = await app.Sheets(1).FieldDescriptors
  const fieldDesc = fieldDescs.FieldDescriptor('Phone','电话')
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const fieldDescs = Application.Sheets(1).FieldDescriptors
  const fieldDesc = fieldDescs.FieldDescriptor('Phone','电话')
}
main()
```
 
# 412 API文档 / API / FieldDescriptors / Item

本页内容

- 说明
- 语法
- 参数
- 返回值
- jsApi 示例
- 脚本编辑器 示例

# FieldDescriptors.Item(方法) ​

## 说明 ​

返回字段的信息，可以通过 FieldDescriptors.Item() 返回某个字段的信息，也可以简化使用FieldDescriptors(Index) 来返回字段信息。

## 语法 ​

表达式.Item(Index)

表达式:FieldDescriptors

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Index | 否 | string/number | Index为number类型时表示当前字段的索引，Index为字段串类型的时候表示字段的Id或@字段名称 |

## 返回值 ​

FieldDescriptor

## jsApi 示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const field = await app.Sheets(1).FieldDescriptors(2)
  console.log(await field.Name)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const field = Application.Sheets(1).FieldDescriptors(2)
  console.log(await field.Name)
}
main()
```
 
# 413 API文档 / API / FieldDescriptors / OnCreate

本页内容

- 说明
- 语法
- 参数
- 返回值
- 事件返回数据
- 浏览器环境示例
- 脚本编辑器 示例

# 监听增加字段的事件 ​

FieldDescriptors.OnCreate(方法)

## 说明 ​

为 FieldDescriptors 添加 Create 事件,当添加 FieldDescriptors 时触发。注意在脚本编辑器中使用时，脚本运行结束就会退出运行，这时可能回调无法被正常触发

## 语法 ​

表达式.OnCreate(Callback)

表达式: FieldDescriptors

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Callback | 是 | func | 指定事件发生时的回调方法 ; const eventContext = await FieldDescriptors.OnCreate(()=&gt;{ ... }) |

## 返回值 ​

EventContext

## 事件返回数据 ​

FieldDescriptor

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    let eventContext;
    eventContext = await app.Sheets(1).FieldDescriptors.OnCreate(data => {
        console.log(data);
        // 取消事件监听
        eventContext.Destroy();
    });
    const desc = await app.FieldDescriptor('Rating', '等级字段');
    desc.MaxRating = 2;
    await app.Sheets(1).FieldDescriptors.AddField(desc, 1);
    //这里会执行OnCreate的回调
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    let eventContext;
    eventContext = Application.Sheets(1).FieldDescriptors.OnCreate(data => {
        console.log(data);
        // 取消事件监听
        eventContext.Destroy();
    });
    const desc = Application.FieldDescriptor('Rating', '等级字段');
    desc.MaxRating = 2;
    Application.Sheets(1).FieldDescriptors.AddField(desc, 1);
    //这里会执行OnCreate的回调
}
main();
```
 
# 414 API文档 / API / Fields / Fields对象

本页内容

# Fields (对象) ​

## 说明 ​

网格视图列的集合，可以访问视图的列属性，通过Item方法访问列时，注意列如果是隐藏了就无法通过索引获取，只能通过名称或者字段ID来获取

## 方法 ​

- [Item](/documents/app-integration-dev/guide/dbsheet/Api/Fields_Item.html)

## 属性 ​

- [Count](/documents/app-integration-dev/guide/dbsheet/Api/Fields_Count.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const fields = await app.Sheets(1).Views(1).Fields
 console.log(await fields.Item(1).Name)
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const fields = Application.Sheets(1).Views(1).Fields
  console.log(fields.Item(1).Name)
}
main()
```
 
# 415 API文档 / API / Fields / Count

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# FieldDescriptors.Count(属性) ​

## 说明 ​

可读写

字段描述的数量

## 返回值 ​

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
   const app = instance.Application;
   const count = await app.Sheets(1).FieldDescriptors.Count
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const count = Application.Sheets(1).FieldDescriptors.Count
}
main()
```
 
# 416 API文档 / API / Fields / Item

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# Fields.Item(方法) ​

## 说明 ​

返回视图字段的信息，可以通过 Fields.Item() 返回某个字段的信息，也可以简化使用Fields(Index) 来返回字段信息。通过Item方法访问列时，注意列如果是隐藏了就无法通过索引获取，只能通过名称或者字段ID来获取。

## 语法 ​

表达式.Item(Index)

表达式:Fields

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Index | 否 | string/number | Index为number类型时表示当前字段的索引，Index为字段串类型的时候, 如果是以@开头则表示是字段名，否则表示字段的Id |

## 返回值 ​

[Field](/documents/app-integration-dev/guide/dbsheet/Api/Field.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const field = await app.Sheets(1).Views(1).Fields.Item(1)
  console.log(await field.Id)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const field = Application.Sheets(1).Views(1).Fields.Item(1)
  console.log(field)
}
main()
```
 
# 417 API文档 / API / FormulaField / FormulaField对象

本页内容

- 说明
- 属性
- 浏览器环境示例
- 脚本编辑器 示例

# FormulaField (对象) ​

## 说明 ​

FormulaField 公式字段的属性，修改属性之后需要 FieldDescriptor 调用Apply()方法使得修改生效。如果非公式字段，无法设置相关属性

## 属性 ​

- [Formula](/documents/app-integration-dev/guide/dbsheet/Api/FormulaField_Formula.html)
- [ValueType](/documents/app-integration-dev/guide/dbsheet/Api/FormulaField_ValueType.html)
- [IsShowPercentAsProgress](/documents/app-integration-dev/guide/dbsheet/Api/FormulaField_IsShowPercentAsProgress.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    // 新建一个公式字段
    await instance.ready();
    const app = instance.Application;
    const fieldDescriptors = await app.Sheets("数据表").FieldDescriptors
    const desc = await app.FieldDescriptor("Formula","公式字段")
    const prop = await desc.Formula
    prop.Formula = '=[@日期]+[@日期]'
    prop.IsShowPercentAsProgress = true
    const result = await fieldDescriptors.AddField(desc)
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
     // 新建一个公式字段
    const fieldDescriptors = Application.Sheets("数据表").FieldDescriptors
    const desc = FieldDescriptor("Formula","公式字段")
    const prop = desc.Formula
    prop.Formula = '=[@日期]+[@日期]'
    prop.IsShowPercentAsProgress = true
    const result = fieldDescriptors.AddField(desc)
 }
main()
```
 
# 418 API文档 / API / FormulaField / Formula

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置公式的文本表达式 ​

FieldDescriptor.Formula(属性)

## 说明 ​

可读写

公式字段 返回公式的文本表达式

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const Application = instance.Application
  const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@公式")
  const prop = await fieldDescriptor.Formula
  prop.Formula = '=[@日期]+[@日期]'
  fieldDescriptor.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@公式")
  const prop = fieldDescriptor.Formula
  prop.Formula = '=[@日期]+[@日期]'
  fieldDescriptor.Apply()
}
main()
```
 
# 419 API文档 / API / FormulaField / IsShowPercentAsProgress

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# FormulaField.IsShowPercentAsProgress(属性) ​

## 说明 ​

可读写

公式字段 “将百分比显示为进度” 属性, 只有当字段属性ValueType 设置为

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const Application = instance.Application
  const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@公式")
  const prop = await fieldDescriptor.Formula
  prop.IsShowPercentAsProgress = true
  fieldDescriptor.Apply()
  console.log(await prop.IsShowPercentAsProgress)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@公式")
  const prop = fieldDescriptor.Formula
  prop.IsShowPercentAsProgress = true
  fieldDescriptor.Apply()
}
main()
```
 
# 420 API文档 / API / FormulaField / ValueType

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# FormulaField.ValueType(属性) ​

## 说明 ​

只读 公式字段返回的格式，根据公式的结果返回相关的类型

## 返回值 ​

[DbFieldValueType](/documents/app-integration-dev/guide/dbsheet/Api/Enum_DbFieldValueType.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const Application = instance.Application
  const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@公式")
  const prop = await fieldDescriptor.Formula
  console.log(await prop.ValueType)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@公式")
  const prop = fieldDescriptor.Formula
  console.log(prop.ValueType)
}
main()
```
 
# 421 API文档 / API / LinkField / LinkField对象

本页内容

- 说明
- 属性
- 浏览器环境示例
- 脚本编辑器 示例

# LinkField (对象) ​

## 说明 ​

LinkField 关联字段的属性，修改属性之后需要 FieldDescriptor 调用Apply()方法使得修改生效。如果非关联字段，无法设置相关属性

## 属性 ​

- [LinkSheet](/documents/app-integration-dev/guide/dbsheet/Api/LinkField_LinkSheet.html)
- [IsAutoLink](/documents/app-integration-dev/guide/dbsheet/Api/LinkField_IsAutoLink.html)
- [IsSupportMultiLinks](/documents/app-integration-dev/guide/dbsheet/Api/LinkField_IsSupportMultiLinks.html)
- [LinkView](/documents/app-integration-dev/guide/dbsheet/Api/LinkField_LinkView.html)
- [AutoLinkGroups](/documents/app-integration-dev/guide/dbsheet/Api/LinkField_AutoLinkGroups.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const Application = instance.Application
  const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@查找引用")
  const linkFieldId = await Application.Sheets("数据表").FieldDescriptors("@单向关联")
  const prop = await fieldDescriptor.Lookup
  prop.LinkFieldId = linkFieldId
  fieldDescriptor.Apply()
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const fieldDescriptor = Sheets("数据表").FieldDescriptors("@查找引用")
  const linkFieldId = Sheets("数据表").FieldDescriptors("@单向关联")
  const prop = fieldDescriptor.Lookup
  prop.LinkFieldId = linkFieldId
  fieldDescriptor.Apply()
 }
main()
```
 
# 422 API文档 / API / LinkField / AutoLinkGroups

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置关系组集合 ​

FieldDescriptor.AutoLinkGroups(属性)

## 说明 ​

可读写

设置或读取 关联字段的关系组集合，如果有多个关系组，则这些关系组是或的关系

![alt text](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/FieldDescriptor_AutoLinkGroups.Dg5N6jJT.png)

## 返回值 ​

AutoLinkGroups

## 浏览器环境示例 ​

javascript

```javascript
// 设置数据表的关联字段的匹配条件
// 匹配到数据表2：数据表2的状态字段 = 数据表的文本字段，如下设置：
async function example() {
  await instance.ready();
  const Application = instance.Application;
  // 数据表
  const sheet = Application.Sheets(1)
  // 数据表的字段
  const fieldId = await sheet.FieldId("文本")
  // 数据表的关联字段
  const linkField = await sheet.FieldDescriptors("@关联")
  const prop = linkField.Link
  const linkGroups = await Application.AutoLinkGroups()
  const group = linkGroups.Add()
  const conditions = group.Conditions
  // 关联的数据表2
  const linkSheet = await Application.Sheets(2)
  // 关联数据表2的字段
  const linkSheet_fieldId = await linkSheet.FieldId("状态")
  // 生成匹配条件
  conditions.Add(linkSheet_fieldId, [fieldId], "Field", "Equals")
  // 设置关联字段的匹配条件
  prop.AutoLinkGroups = linkGroups
  // 设置自动关联
  prop.IsAutoLink = true
  linkField.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main(){
  const sheet = Application.Sheets(1)
  const fieldId = sheet.FieldId("状态")
  // 获取关联字段
  const linkField = sheet.FieldDescriptors("@关联")
  const prop = linkField.Link
  const linkGroups = Application.AutoLinkGroups()
  const conditions = linkGroups.Conditions
  // 关联的数据表2
  const linkSheet = Application.Sheets(2)
  // 关联数据表2的字段
  const linkSheet_fieldId = linkSheet.FieldId("状态")
  // 生成匹配条件
  conditions.Add(linkSheet_fieldId, [fieldId], "Field", "Equals")
  // 设置关联字段的匹配条件
  prop.AutoLinkGroups = linkGroups
  // 设置自动关联
  prop.IsAutoLink = true
  linkField.Apply()
}
main()
```
 
# 423 API文档 / API / LinkField / IsAutoLink

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 自动关联 ​

LinkField.IsAutoLink(属性)

## 说明 ​

可读写

设置或读取 单向关联字段 或 双向关联字段 是否自动关联

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const field = await app.Sheets(1).FieldDescriptors(2)
    const prop = await field.Link
    console.log(await prop.IsAutoLink)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const field = Application.Sheets(1).FieldDescriptors(2)
    const prop = field.Link
    console.log(prop.IsAutoLink)
}
main()
```
 
# 424 API文档 / API / LinkField / IsSupportMultiLinks

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 关联多条记录 ​

LinkField.IsSupportMultiLinks(属性)

## 说明 ​

可读写

设置或读取 单向关联字段或双向关联字段 是否可以设置允许关联多条记录

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const field = await app.Sheets(1).FieldDescriptors(2)
    const link = await field.Link
    field.IsSupportMultiLinks = true
	field.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const field = Application.Sheets(1).FieldDescriptors(2)
    const link = field.Link
    link.IsSupportMultiLinks = true
	field.Apply() 
}
main()
```
 
# 425 API文档 / API / LinkField / LinkSheet

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 关联表格ID ​

FieldDescriptor.LinkSheet(属性)

## 说明 ​

可读写

设置或读取单向关联或双向关联字段的关联表格的Id

## 返回值 ​

Number

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    // 插入单向关联字段
    const Application = instance.Application
    const sheet = await Application.Sheets(1)
    const linkField = await Application.FieldDescriptor("OneWayLink","单向关联")
    const prop = await linkField.Link
    prop.LinkSheet = await sheet.StId
    prop.IsAutoLink = false
    sheet.FieldDescriptors.AddField(linkField)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const sheet = Application.Sheets(1)
    const linkField = Application.FieldDescriptor("OneWayLink","单向关联")
    const prop = linkField.Link
    prop.LinkSheet = sheet.StId
    prop.IsAutoLink = false
    sheet.FieldDescriptors.AddField(linkField)
}
main()
```
 
# 426 API文档 / API / LinkField / LinkView

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置关联视图ID ​

FieldDescriptor.LinkView(属性)

## 说明 ​

可读写

设置或读取单向关联或双向关联字段的关联视图的id

![alt text](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/FieldDescriptor_LinkView.BeJbQ7Ke.png)

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const field = await app.Sheets(1).FieldDescriptors(2)
    const prop = await field.Link
    prop.LinkView = await app.ActiveView.Id
    field.Apply()
    console.log(await prop.LinkView)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const field = Application.Sheets(1).FieldDescriptors(2)
    const prop = field.Link
    prop.LinkView = ActiveView.Id
    field.Apply()
    console.log(prop.LinkView)
}
main()
```
 
# 427 API文档 / API / LookupField / LookupField对象

本页内容

- 说明
- 属性
- 浏览器环境示例
- 脚本编辑器 示例

# LookupField (对象) ​

## 说明 ​

LookupField 引用/查找引用/统计 字段的属性，修改属性之后需要 FieldDescriptor 调用Apply()方法使得修改生效。如果非 引用/查找引用/统计 字段，无法设置相关属性

## 属性 ​

- [LinkFieldId](/documents/app-integration-dev/guide/dbsheet/Api/LookupField_LinkFieldId.html)
- [LookupFieldId](/documents/app-integration-dev/guide/dbsheet/Api/LookupField_LookupFieldId.html)
- [LookupType](/documents/app-integration-dev/guide/dbsheet/Api/LookupField_LookupType.html)
- [LookupSheetId](/documents/app-integration-dev/guide/dbsheet/Api/LookupField_LookupSheetId.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const Application = instance.Application
  const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@查找引用")
  const linkFieldId = await Application.Sheets("数据表").FieldDescriptors("@单向关联")
  const prop = await fieldDescriptor.Lookup
  prop.LinkFieldId = linkFieldId
  fieldDescriptor.Apply()
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const fieldDescriptor = Sheets("数据表").FieldDescriptors("@查找引用")
  const linkFieldId = Sheets("数据表").FieldDescriptors("@单向关联")
  const prop = fieldDescriptor.Lookup
  prop.LinkFieldId = linkFieldId
  fieldDescriptor.Apply()
 }
main()
```
 
# 428 API文档 / API / LookupField / LinkFieldId

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# FieldDescriptor.LinkFieldId(属性) ​

## 说明 ​

可读写

引用内容 设置为 关联的字段Id 以联接数据表

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const Application = instance.Application
  const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@查找引用")
  const linkFieldId = await Application.Sheets("数据表").FieldDescriptors("@单向关联")
  const prop = await fieldDescriptor.Lookup
  prop.LinkFieldId = linkFieldId
  fieldDescriptor.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const fieldDescriptor = Sheets("数据表").FieldDescriptors("@查找引用")
  const linkFieldId = Sheets("数据表").FieldDescriptors("@单向关联")
  const prop = fieldDescriptor.Lookup
  prop.LinkFieldId = linkFieldId
  fieldDescriptor.Apply()
}
main()
```
 
# 429 API文档 / API / LookupField / LookupFieldId

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# LookupField.LookupFieldId(属性) ​

## 说明 ​

可读写

引用内容里 引用字段的Id，指向的是引用的数据表的字段（不一定是当前数据表）,需要先从关联的表中读取字段Id。

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const Application = instance.Application
  const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@查找引用")
  const prop = await fieldDescriptor.Lookup
  const lookUpSheet = await Application.Sheets.ItemById(prop.LookupSheetId)
  const lookupField = await lookUpSheet.FieldDescriptors(prop.LookupFieldId)
  console.log(await lookupField.Name)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@查找引用")
  const prop = fieldDescriptor.Lookup
  const lookUpSheet = Application.Sheets.ItemById(prop.LookupSheetId)
  const lookupField = lookUpSheet.FieldDescriptors(prop.LookupFieldId)
  console.log(lookupField.Name)
}
main()
```
 
# 430 API文档 / API / LookupField / LookupSheetId

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# FieldDescriptor.LookupSheetId(属性) ​

## 说明 ​

可读写

引用字段里 引用sheet的Id，要通过Sheet Id获取Sheet的实例可以通过 Application.Sheets.ItemById(id)

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const Application = instance.Application
  const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@查找引用")
  const prop = await fieldDescriptor.Lookup
  const lookUpSheet = await Application.Sheets.ItemById(prop.LookupSheetId)
  const lookupField = await lookUpSheet.FieldDescriptors(prop.LookupFieldId)
  console.log(await lookupField.Name)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@查找引用")
  const prop = fieldDescriptor.Lookup
  const lookUpSheet = Application.Sheets.ItemById(prop.LookupSheetId)
  const lookupField = lookUpSheet.FieldDescriptors(prop.LookupFieldId)
  console.log(lookupField.Name)
}
main()
```
 
# 431 API文档 / API / LookupField / LookupType

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# LookupField.LookupType(属性) ​

## 说明 ​

可读写

引用字段的显示方法,查找引用字段不允许配置，只能使用Origin，统计字段不允许设置Origin

## 返回值 ​

[DbLookupFunction](/documents/app-integration-dev/guide/dbsheet/Api/Enum_DbLookupFunction.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const Application = instance.Application
  const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@查找引用")
  const prop = await fieldDescriptor.Lookup
  console.log(await prop.LookupType)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@查找引用")
  const prop = fieldDescriptor.Lookup
  console.log(prop.LookupType)
}
main()
```
 
# 432 API文档 / API / NumberField / NumberField对象

本页内容

- 说明
- 属性
- 浏览器环境示例
- 脚本编辑器 示例

# NumberField (对象) ​

## 说明 ​

NumberField 数字字段的属性，修改属性之后需要 FieldDescriptor 调用Apply()方法使得修改生效。如果非数字字段，无法设置相关属性

## 属性 ​

- [IsShowThousand](/documents/app-integration-dev/guide/dbsheet/Api/NumberField_IsShowThousand.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const field = await app.Sheets(1).FieldDescriptors("@数字")
  const prop = await field.Number
  prop.IsShowThousand = true
  field.Apply()
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const field = Application.Sheets(1).FieldDescriptors("@数字")
  const prop = field.Number
  prop.IsShowThousand = true
  field.Apply()
 }
main()
```
 
# 433 API文档 / API / NumberField / IsShowThousand

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 是否显示千位符 ​

NumberField.IsShowThousand(属性)

## 说明 ​

可读写

字段类型为Number时，通过此属性可以快捷的设置是否显示千位符

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const field = await app.Sheets(1).FieldDescriptors("@数字")
  const prop = await field.Number
  prop.IsShowThousand = true
  field.Apply()
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const field = Application.Sheets(1).FieldDescriptors("@数字")
  const prop = field.Number
  prop.IsShowThousand = true
  field.Apply()
 }
main()
```
 
# 434 API文档 / API / RatingField / RatingField对象

本页内容

- 说明
- 属性
- 浏览器环境示例
- 脚本编辑器 示例

# RatingField (对象) ​

## 说明 ​

RatingField 等级字段的属性，修改属性之后需要 FieldDescriptor 调用Apply()方法使得修改生效。如果等级字段，无法设置相关属性

## 属性 ​

- [MaxRating](/documents/app-integration-dev/guide/dbsheet/Api/RatingField_MaxRating.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const field = await app.Sheets(1).FieldDescriptors("@数字")
  const prop = await field.Rating
  prop.MaxRating = 1
  field.Apply()
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const field = Application.Sheets(1).FieldDescriptors("@数字")
  const prop = field.Rating
  prop.MaxRating = 2
  field.Apply()
 }
main()
```
 
# 435 API文档 / API / RatingField / MaxRating

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置等级 ​

RatingField.MaxRating(属性)

## 说明 ​

可读写

只对等级字段有效，设置等级字段最大值，取值范围为1-5

## 返回值 ​

Number

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const field = await app.Sheets(1).FieldDescriptors(2)
  const prop = await field.Rating
  prop.MaxRating = 5
  field.Apply()

  console.log(await prop.MaxRating)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const field = Application.Sheets(1).FieldDescriptors(2)
  const prop = await field.Rating
  prop.MaxRating = 5
  field.Apply()

  console.log(prop.MaxRating)
}
main()
```
 
# 436 API文档 / API / SelectField / SelectField对象

本页内容

- 说明
- 属性
- 浏览器环境示例
- 脚本编辑器 示例

# SelectField (对象) ​

## 说明 ​

SelectField 单选项/多选项字段的属性，修改属性之后需要 FieldDescriptor 调用Apply()方法使得修改生效。如果单选项/多选项字段，无法设置相关属性

## 属性 ​

- [Items](/documents/app-integration-dev/guide/dbsheet/Api/SelectField_Items.html)
- [IsAddItemWhenInputting](/documents/app-integration-dev/guide/dbsheet/Api/SelectField_IsAddItemWhenInputting.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const filed = await app.Sheets(1).FieldDescriptors("@状态")
  const prop = await filed.Select
  const Items = []
  Items.push({"value":"未开始233", "colorHex":"#0081C2"})
  prop.Items = Items
  filed.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main(){
  const filed = Application.Sheets(1).FieldDescriptors("@状态")
  const prop = field.Select
  const Items = []
  Items.push({"value":"未开始233", "colorHex":"#0081C2"})
  prop.Items = Items
  filed.Apply()
}
main()
```
 
# 437 API文档 / API / SelectField / IsAddItemWhenInputting

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 允许填写时添加选项 ​

SelectField.IsAddItemWhenInputting(属性)

## 说明 ​

可读写

当字段类型为单选项或多选项时，可以通过设置 IsAddItemWhenInputting 属性设置允许填写时添加选项。当字段类型不是单选项或多选项时，属性无效。

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
   const app = instance.Application;
   const field = await app.Sheets(1).FieldDescriptors(2)
   const prop = await field.Select
   prop.IsAddItemWhenInputting = true
   field.Apply()
   console.log(await prop.IsAddItemWhenInputting)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const field = Application.Sheets(1).FieldDescriptors(2)
   const prop = field.Select
   prop.AllowAddItemWhenInputting = true
   field.Apply()
   console.log(prop.AllowAddItemWhenInputing)
}
main()
```
 
# 438 API文档 / API / SelectField / Items

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置选项 ​

SelectField.Items(属性)

## 说明 ​

可读写

设置多项式字段的可选项

## 返回值 ​

Array

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const filed = await app.Sheets(1).FieldDescriptors("@状态")
  const prop = await filed.Select
  const Items = await prop.Items
  // 加一项
  Items.push({"value":"未开始233", "colorHex":"#0081C2"})
  prop.Items = Items
  filed.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main(){
  const filed = Application.Sheets(1).FieldDescriptors("@状态")
  const prop = field.Select
  const Items = prop.Items
  // 加一项
  Items.push({"value":"未开始233", "colorHex":"#0081C2"})
  prop.Items = Items
  filed.Apply()
}
main()
```
 
# 439 API文档 / API / TextLinkRun / TextLinkRun对象

本页内容

- 说明
- 方法
- 属性
- 浏览器环境示例
- 脚本编辑器 示例

# TextLinkRun (对象) ​

## 说明 ​

TextLinkRun代表文本链接的对象。

## 方法 ​

## 属性 ​

- [Address](/documents/app-integration-dev/guide/dbsheet/Api/TextLinkRun_Address.html)
- [Pos](/documents/app-integration-dev/guide/dbsheet/Api/TextLinkRun_Pos.html)
- [Length](/documents/app-integration-dev/guide/dbsheet/Api/TextLinkRun_Length.html)
- [LinkRunsType](/documents/app-integration-dev/guide/dbsheet/Api/TextLinkRun_LinkRunsType.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const app = instance.Application;
   const recordComment = await app.ActiveView.RecordComments.Item(1)
   recordComment.Add("Hello World")
   const count = await recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = await recordComment.Item(i)
    const textLinkRuns = await comment.TextLinkRuns
    const linkCount =  await textLinkRuns.Count
    for (let j = 1; j <= linkCount; j++) {
     const textLinkRun = await textLinkRuns.Item(j)
     console.log(await textLinkRun.Address)
    }
   }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const recordComment = ActiveView.RecordComments.Item(1)
   recordComment.Add("Hello World")
   const count = recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = recordComment.Item(i)
    const textLinkRuns = comment.TextLinkRuns
    console.log(textLinkRuns.Count)
   }
}
main()
```
 
# 440 API文档 / API / TextLinkRun / Address

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# TextLinkRun.Address(属性) ​

## 说明 ​

只读 返回文本链接的地址

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const app = instance.Application;
   const recordComment = await app.ActiveView.RecordComments.Item(1)
   recordComment.Add("Hello World")
   const count = await recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = await recordComment.Item(i)
    const textLinkRuns = await comment.TextLinkRuns
    const linkCount =  await textLinkRuns.Count
    for (let j = 1; j <= linkCount; j++) {
     const textLinkRun = await textLinkRuns.Item(j)
     console.log(await textLinkRun.Address)
    }
   }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const recordComment = ActiveView.RecordComments.Item(1)
   const count = recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = recordComment.Item(i)
    const textLinkRuns = comment.TextLinkRuns
    const linkCount =  textLinkRuns.Count
    for (let j = 1; j <= linkCount; j++) {
     const textLinkRun = textLinkRuns.Item(j)
     console.log(textLinkRun.Address)
    }
   }
}
main()
```
 
# 441 API文档 / API / TextLinkRun / Length

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# TextLinkRun.Length(属性) ​

## 说明 ​

只读 返回的文本链接的长度

## 返回值 ​

Number

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const app = instance.Application;
   const recordComment = await app.ActiveView.RecordComments.Item(1)
   recordComment.Add("Hello World")
   const count = await recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = await recordComment.Item(i)
    const textLinkRuns = await comment.TextLinkRuns
    const linkCount =  await textLinkRuns.Count
    for (let j = 1; j <= linkCount; j++) {
     const textLinkRun = await textLinkRuns.Item(j)
     console.log(await textLinkRun.Length)
    }
   }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const recordComment = ActiveView.RecordComments.Item(1)
   const count = recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = recordComment.Item(i)
    const textLinkRuns = comment.TextLinkRuns
    const linkCount =  textLinkRuns.Count
    for (let j = 1; j <= linkCount; j++) {
     const textLinkRun = textLinkRuns.Item(j)
     console.log(textLinkRun.Length)
    }
   }
}
main()
```
 
# 442 API文档 / API / TextLinkRun / LinkRunsType

本页内容

# TextLinkRun.LinkRunsType(属性) ​

## 说明 ​

只读 返回文本链接的类型

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const app = instance.Application;
   const recordComment = await app.ActiveView.RecordComments.Item(1)
   recordComment.Add("Hello World")
   const count = await recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = await recordComment.Item(i)
    const textLinkRuns = await comment.TextLinkRuns
    const linkCount =  await textLinkRuns.Count
    for (let j = 1; j <= linkCount; j++) {
     const textLinkRun = await textLinkRuns.Item(j)
     console.log(await textLinkRun.LinkRunsType)
    }
   }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const recordComment = ActiveView.RecordComments.Item(1)
   const count = recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = recordComment.Item(i)
    const textLinkRuns = comment.TextLinkRuns
    const linkCount =  textLinkRuns.Count
    for (let j = 1; j <= linkCount; j++) {
     const textLinkRun = textLinkRuns.Item(j)
     console.log(textLinkRun.LinkRunsType)
    }
   }
}
main()
```
 
# 443 API文档 / API / TextLinkRun / Pos

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# TextLinkRun.Pos(属性) ​

## 说明 ​

只读 返回链接的开始位置，以字符为单位。

## 返回值 ​

Number

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const app = instance.Application;
   const recordComment = await app.ActiveView.RecordComments.Item(1)
   recordComment.Add("Hello World")
   const count = await recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = await recordComment.Item(i)
    const textLinkRuns = await comment.TextLinkRuns
    const linkCount =  await textLinkRuns.Count
    for (let j = 1; j <= linkCount; j++) {
     const textLinkRun = await textLinkRuns.Item(j)
     console.log(await textLinkRun.Pos)
    }
   }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const recordComment = ActiveView.RecordComments.Item(1)
   const count = recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = recordComment.Item(i)
    const textLinkRuns = comment.TextLinkRuns
    const linkCount =  textLinkRuns.Count
    for (let j = 1; j <= linkCount; j++) {
     const textLinkRun = textLinkRuns.Item(j)
     console.log(textLinkRun.Pos)
    }
   }
}
main()
```
 
# 444 API文档 / API / TextLinkRuns / TextLinkRuns对象

本页内容

# TextLinkRuns (对象) ​

## 说明 ​

TextLinkRuns 对象代表文本链接的集合。

## 方法 ​

- [Item](/documents/app-integration-dev/guide/dbsheet/Api/TextLinkRuns_Item.html)

## 属性 ​

- [Count](/documents/app-integration-dev/guide/dbsheet/Api/TextLinkRuns_Count.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const app = instance.Application;
   const recordComment = await app.ActiveView.RecordComments.Item(1)
   recordComment.Add("Hello World")
   const count = await recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = await recordComment.Item(i)
    const textLinkRuns = await comment.TextLinkRuns
    const linkCount =  await textLinkRuns.Count
    for (let j = 1; j <= linkCount; j++) {
     const textLinkRun = await textLinkRuns.Item(j)
     console.log(await textLinkRun.Address)
    }
   }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const recordComment = ActiveView.RecordComments.Item(1)
   recordComment.Add("Hello World")
   const count = recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = recordComment.Item(i)
    const textLinkRuns = comment.TextLinkRuns
    console.log(textLinkRuns.Count)
   }
}
main()
```
 
# 445 API文档 / API / TextLinkRuns / Count

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# TextLinkRuns.Count(属性) ​

## 说明 ​

只读 返回的记录里包含TextLinkRun的数量

## 返回值 ​

Number

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const app = instance.Application;
   const recordComment = await app.ActiveView.RecordComments.Item(1)
   recordComment.Add("Hello World")
   const count = await recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = await recordComment.Item(i)
    const textLinkRuns = await comment.TextLinkRuns
    const linkCount =  await textLinkRuns.Count
    for (let j = 1; j <= linkCount; j++) {
     const textLinkRun = await textLinkRuns.Item(j)
     console.log(await textLinkRun.Address)
    }
   }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const recordComment = ActiveView.RecordComments.Item(1)
   const count = recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = recordComment.Item(i)
    const textLinkRuns = comment.TextLinkRuns
    const linkCount =  textLinkRuns.Count
    for (let j = 1; j <= linkCount; j++) {
     const textLinkRun = textLinkRuns.Item(j)
     console.log(textLinkRun.Address)
    }
   }
}
main()
```
 
# 446 API文档 / API / TextLinkRuns / Item

本页内容

- 说明
- 语法
- 参数
- 返回值
- jsApi 示例
- 脚本编辑器 示例

# TextLinkRuns.Item(方法) ​

## 说明 ​

获取指定索引位置或评论ID的记录

## 语法 ​

表达式.Item(Index)

表达式:TextLinkRuns

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Index | 否 | number | 传入number时索引从1开始 |

## 返回值 ​

[TextLinkRun](/documents/app-integration-dev/guide/dbsheet/Api/TextLinkRun.html)

## jsApi 示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const app = instance.Application;
   const recordComment = await app.ActiveView.RecordComments(1)
   const count = recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = await recordComment.Item(i)
    console.log(await comment.Text)
   }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const recordComment = ActiveView.RecordComments(1)
   const count = recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = recordComment.Item(i)
    console.log(comment.Text)
   }
}
main()```
```
 
# 447 API文档 / API / UrlField / UrlField对象

本页内容

- 说明
- 属性
- 浏览器环境示例
- 脚本编辑器 示例

# UrlField (对象) ​

## 说明 ​

UrlField 超链接字段的属性，修改属性之后需要 FieldDescriptor 调用Apply()方法使得修改生效。如果非超链接字段，无法设置相关属性

## 属性 ​

- [HyperLinkText](/documents/app-integration-dev/guide/dbsheet/Api/UrlField_HyperLinkText.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const field = await app.Sheets(1).FieldDescriptors("@超链接")
  const prop = await field.Url
  prop.HyperLinkText = "Go"
  field.Apply()
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const field = Application.Sheets(1).FieldDescriptors("@超链接")
  const prop = field.Url
  prop.HyperLinkText = "Go"
  field.Apply()
 }
main()
```
 
# 448 API文档 / API / UrlField / HyperLinkText

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置显示样式 ​

UrlField.HyperLinkText(属性)

## 说明 ​

可读写

只对超链接字段有效，超链接字段设置显示的文本，如果设置了 HyperLinkText ，超链接字段显示样式就会以按钮形式显示

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const field = await app.Sheets(1).FieldDescriptors("@超链接")
  const prop = await field.Url
  // 以按钮形式显示 "Go"
  prop.HyperLinkText = "Go"
  field.Apply()

  // 以超链接形式显示
  prop.HyperLinkText = ""
  field.Apply()
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const field = Application.Sheets(1).FieldDescriptors("@超链接")
  const prop = field.Url
  // 以按钮形式显示 "Go"
  prop.HyperLinkText = "Go"
  field.Apply()

  // 以超链接形式显示
  prop.HyperLinkText = ""
  field.Apply()
}
main()
```
 
# 449 API文档 / API / WatchedField / WatchedField对象

本页内容

- 说明
- 属性
- 浏览器环境示例
- 脚本编辑器 示例

# WatchedField (对象) ​

## 说明 ​

WatchedField 最后修改人/最后修改时间 字段的属性，修改属性之后需要 FieldDescriptor 调用Apply()方法使得修改生效。如果最后修改人/最后修改时间，无法设置相关属性

## 属性 ​

- [IsWatchedAll](/documents/app-integration-dev/guide/dbsheet/Api/WatchedField_IsWatchedAll.html)
- [WatchedFields](/documents/app-integration-dev/guide/dbsheet/Api/WatchedField_WatchedFields.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const Application = instance.Application
    const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@最后修改时间")
    const prop = await fieldDescriptor.Watch
    prop.IsWatchedAll = true
    fieldDescriptor.Apply()
    console.log(await prop.IsWatchedAll)
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@最后修改时间")
    const prop = fieldDescriptor.Watch
    prop.IsWatchedAll = true
    fieldDescriptor.Apply()
 }
main()
```
 
# 450 API文档 / API / WatchedField / IsWatchedAll

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 监听所有字段 ​

WatchedField.IsWatchedAll(属性)

## 说明 ​

可读写

最后修改人/最后修改时间，是否监听 所有字段的修改

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const Application = instance.Application
    const fieldDescriptor = await Application.Sheets("数据表").FieldDescriptors("@最后修改时间")
    const prop = await fieldDescriptor.Watch
    prop.IsWatchedAll = true
    fieldDescriptor.Apply()
    console.log(await prop.IsWatchedAll)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@最后修改时间")
    const prop = fieldDescriptor.Watch
    prop.IsWatchedAll = true
    fieldDescriptor.Apply()
}
main()
```
 
# 451 API文档 / API / WatchedField / WatchedFields

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 监听某些字段 ​

WatchedField.WatchedFields(属性)

## 说明 ​

可读写

最后修改人/最后修改时间，监听某些字段的修改情况, 如果需要监听某些特定的字段，需要将属性 IsWatchedAll 设置为false, 否则这个设置不生效

## 返回值 ​

Array

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const Application = instance.Application
    const fieldDescriptors = await Application.Sheets("数据表").FieldDescriptors
    const fieldDescriptor = await fieldDescriptors.Item("@最后修改时间")
    const prop = await fieldDescriptor.Watch
    const watchId = await fieldDescriptors.Item("@文本").Id
    prop.IsWatchedAll = false
    prop.WatchedFields = [watchId]
    fieldDescriptor.Apply()
    console.log(await prop.WatchedFields)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const fieldDescriptor = Application.Sheets("数据表").FieldDescriptors("@最后修改时间")
    const watchId = await Application.Sheets("数据表").FieldDescriptors("@文本").Id
    const prop = fieldDescriptor.Watch
    prop.IsWatchedAll = false
    prop.WatchedFields = [watchId]
    fieldDescriptor.Apply()
}
main()
```
 
# 452 API文档 / API / Sort / Sort对象

本页内容

- 说明
- 方法
- 属性
- 浏览器环境示例
- 脚本编辑器 示例

# Sort (对象) ​

## 说明 ​

单条排序记录

## 方法 ​

- [Delete](/documents/app-integration-dev/guide/dbsheet/Api/Sort_Delete.html)

## 属性 ​

- [IsAscending](/documents/app-integration-dev/guide/dbsheet/Api/Sort_IsAscending.html)
- [Field](/documents/app-integration-dev/guide/dbsheet/Api/Sort_Field.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const sorts = await app.Sheets(1).Views(1).Sorts;
    const sort = await sorts(1);
    sort.Delete()
    //const sort = sorts('B');
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const sorts = Application.Sheets(1).Views(1).Sorts;
    const sort = sorts(1);
    sort.Delete()
    //const sort = sorts('B');
}
main()
```
 
# 453 API文档 / API / Sort / Field

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# Sort.Field(属性) ​

## 说明 ​

可读写

排序字段

## 返回值 ​

Field

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const sorts = await app.Sheets(1).Views(1).Sorts;
    const field = sorts(1).Field;

    // 设置排序条件字段
    sorts(1).Field = '@数字';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const sorts = Application.Sheets(1).Views(1).Sorts;
    const field = sorts(1).Field;

    // 设置排序条件字段
    sorts(1).Field = '@数字';
}
main()
```
 
# 454 API文档 / API / Sort / IsAscending

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置排序升序属性 ​

Sort.IsAscending(属性)

## 说明 ​

可读写

设置排序升序或者降序

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
// 获取IsAscending属性
async function example() {
    await instance.ready();
    const app = instance.Application;
    const sorts = await app.Sheets(1).Views(1).Sorts;
    const isAscending = sorts(1).IsAscending;
}

// 设置IsAscending属性
async function example() {
    await instance.ready();
    const app = instance.Application;
    app.Sheets(1).Views(1).Sorts(1).IsAscending = false;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const sorts = Application.Sheets(1).Views(1).Sorts;
    const isAscending = sorts(1).IsAscending;
    isAscending = false;
}
main()
```
 
# 455 API文档 / API / Sort / Delete

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 删除排序条件 ​

Sort.Remove(方法)

## 说明 ​

删除单个排序条件

## 语法 ​

表达式.Delete()

表达式:Sort

## 参数 ​

无参数

## 返回值 ​

[ApiResult](/documents/app-integration-dev/guide/dbsheet/Api/ApiResult.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const sorts = await app.Sheets(1).Views(1).Sorts;
    const res = await sorts(1).Delete();
    if (res.Code === 0) {
        console.log("成功删除")
    } else {
        console.error("删除错误" + res.Message)
    }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const sorts = Application.Sheets(1).Views(1).Sorts;
    const res = sorts(1).Delete();
    if (res.Code === 0) {
        console.log("成功删除")
    } else {
        console.error("删除错误" + res.Message)
    }
}
main()
```
 
# 456 API文档 / API / Sorts / Sorts对象

本页内容

- 说明
- 方法
- 属性
- 浏览器环境示例
- 脚本编辑器 示例

# Sorts (对象) ​

## 说明 ​

当前视图下的排序列表

## 方法 ​

- [Item](/documents/app-integration-dev/guide/dbsheet/Api/Sorts_Item.html)
- [Add](/documents/app-integration-dev/guide/dbsheet/Api/Sorts_Add.html)

## 属性 ​

- [Count](/documents/app-integration-dev/guide/dbsheet/Api/Sorts_Count.html)
- [IsAuto](/documents/app-integration-dev/guide/dbsheet/Api/Sorts_IsAuto.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const sorts = await app.Sheets(1).Views(1).Sorts;
    console.log(await sorts.Count)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const sorts = Application.Sheets(1).Views(1).Sorts;
  console.log(sorts.Count)
}
main()
```
 
# 457 API文档 / API / Sorts / Count

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# Sorts.Count(属性) ​

## 说明 ​

可读 返回排序列表的个数

## 返回值 ​

Number

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const sorts = await app.Sheets(1).Views(1).Sorts;
    const count = sorts.Count;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const sorts = Application.Sheets(1).Views(1).Sorts;
    const count = sorts.Count;
}
main()
```
 
# 458 API文档 / API / Sorts / IsAuto

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# Sorts.IsAuto(属性) ​

## 说明 ​

可读写

自动排序属性

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
// 获取自动排序属性
async function example() {
    await instance.ready();
    const app = instance.Application;
    const sorts = await app.Sheets(1).Views(1).Sorts;
    const isAuto = sorts.IsAuto;
}

// 设置自动排序属性
async function example() {
    await instance.ready();
    const app = instance.Application;
    app.Sheets(1).Views(1).Sorts.isAuto = false;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const sorts = Application.Sheets(1).Views(1).Sorts;
    const isAuto = sorts.IsAuto;
    isAuto = false;
}
main()
```
 
# 459 API文档 / API / Sorts / Add

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 添加排序条件 ​

Sorts.Add(方法)

## 说明 ​

添加排序

## 语法 ​

表达式.Add(Field,IsAscending)

表达式: Sorts

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Field | 是 | number/string | 新增排序字段索引/新增排序字段 ID/新增排序字段名(名称要以@字符作为开始) |
| IsAscending | 否 | boolean | 是否为升序 |

## 返回值 ​

Sort

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const sorts = await app.Sheets(1).Views(1).Sorts;
    const res = sorts.Add(1);
    //const res = sorts.Add('B');
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const sorts = Application.Sheets(1).Views(1).Sorts;
    const res = sorts.Add(1);
    //const res = sorts.Add('B');
}
main()
```
 
# 460 API文档 / API / Sorts / ChangeOrder

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 移动排序条件 ​

Sorts.ChangeOrder(方法)

## 说明 ​

移动排序条件(设置排序优先级)

## 语法 ​

表达式.ChangeOrder(FromField, BeforeField, AfterField)

表达式: Sorts

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| FromField | 是 | [string] | 要移动的排序字段的字段ID/要移动的排序字段的字段名称(名称要以@字符作为开始) |
| BeforeField | 否 | [string] | 目标位置前的排序字段ID/目标位置前的排序字段名称(名称要以@字符作为开始) |
| AfterField | 否 | [string] | 目标位置后的排序字段ID/目标位置后的排序字段名称(名称要以@字符作为开始) |

FromField、BeforeField和AfterField必须都是已设置的排序条件字段，BeforeField和AfterField至少需要传入一个，如果BeforeField和AfterField同时存在以BeforeField作为应用参数

比如表格视图中已设置的排序条件在排序面板中从上到下依次为【公式，日期，名称，数量】 现在想将名称这条排序条件移动到日期的前面，结果变为【公式，名称，日期，数量】，就可以用以下方式实现

```javascript
await WPSOpenApi.Application.Sheets(1).Views(1).Sorts.ChangeOrder('@名称', '@日期')
```

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const sorts = await app.Sheets(1).Views(1).Sorts;
    // 将公式排序条件移动到日期排序条件的后面
    const res = await sorts.ChangeOrder('@公式', undefined, '@日期');
    if (res.Code === 0) {
        console.log("设置排序优先级成功")
    } else {
        console.error("设置排序优先级失败" + res.Message)
    }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const sorts = Application.Sheets(1).Views(1).Sorts;
    // 将公式排序条件移动到日期排序条件的后面
    const res = sorts.ChangeOrder('@公式', undefined, '@日期');
    if (res.Code === 0) {
        console.log("设置排序优先级成功")
    } else {
        console.error("设置排序优先级失败" + res.Message)
    }
}
main()
```
 
# 461 API文档 / API / Sorts / Item

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# Sorts.Item(方法) ​

## 说明 ​

通过索引位置或者字段 ID 获取单条排序记录

## 语法 ​

表达式.Item(Index)

表达式:Sorts

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Index | 是 | number/string | 索引从 1 开始/字段 ID |

## 返回值 ​

Sort

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const sorts = await app.Sheets(1).Views(1).Sorts;
    const sort = sorts.Item(1);
    //const sort = sorts.Item('B');
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const sorts = Application.Sheets(1).Views(1).Sorts;
    const sort = sorts.Item(1);
    //const sort = sorts.Item('B');
}
main()
```
 
# 462 API文档 / API / Criteria / Criteria对象

本页内容

- 说明
- 属性
- 浏览器环境示例
- 脚本编辑器示例

# Criteria (对象) ​

## 说明 ​

筛选条件项

## 属性 ​

- [Field](/documents/app-integration-dev/guide/dbsheet/Api/Criteria_Field.html)
- [Op](/documents/app-integration-dev/guide/dbsheet/Api/Criteria_Op.html)
- [Values](/documents/app-integration-dev/guide/dbsheet/Api/Criteria_Value.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const filters = await app.Sheets(1).Views(1).Filters;
 const criteria = filters(1).Criteria;
 console.log(criteria.Field)
 console.log(criteria.Op)
 console.log(criteria.Values)
 }
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 const filters = Application.Sheets(1).Views(1).Filters
 const result = filters(1).Criteria
 console.log(result.Field)
 console.log(result.Op)
 console.log(result.Values)
 }
main()
```
 
# 463 API文档 / API / Criteria / Field

本页内容

- 说明
- 浏览器环境示例
- 脚本编辑器示例

# Criteria.Field(属性) ​

## 说明 ​

筛选项字段值：

| 值选项 | 描述 |
| --- | --- |
| columnNumber | 列号，从1开始 |
| fieldId | 字段id |

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   // 添加一个筛选条件
    await instance.ready();
    const app = instance.Application;
    const filters = await app.Sheets(1).Views(1).Filters;
    const criteria = app.Criteria(1, 'Equals', ['1'])
    const filter = await filters.Add(criteria);

    // 读取添加的筛选条件规则
    const criteria = await app.Sheets(1).Views(1).Filters(1).Criteria
    console.log(criteria.Op) // "Equals"
    console.log(criteria.Field) // 1
    console.log(criteria.Value) // [{type: 'Text', value: '1'}]
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
     // 添加一个筛选条件
    const filters = Application.Sheets(1).Views(1).Filters;
    const criteria = Application.Criteria(1, 'Equals', ['1'])
    const filter = filters.Add(criteria);

    // 读取添加的筛选条件规则
    const criteria = Application.Sheets(1).Views(1).Filters(1).Criteria
    console.log(criteria.Op) // "Equals"
    console.log(criteria.Field) // 1
    console.log(criteria.Value) // [{type: 'Text', value: '1'}]
}
main()
```
 
# 464 API文档 / API / Criteria / Op

本页内容

- 说明
- 浏览器环境示例
- 脚本编辑器示例

# Criteria.Op(属性) ​

## 说明 ​

筛选项规则（大小写不敏感）：

| 枚举值 | 描述 |
| --- | --- |
| Equals | 等于 |
| NotEqu | 不等于 |
| Greater | 大于 |
| GreaterEqu | 大等于 |
| Less | 小于 |
| LessEqu | 小等于 |
| GreaterEquAndLessEqu | 介于（取等） |
| LessOrGreater | 介于（不取等） |
| BeginWith | 开头是 |
| EndWith | 结尾是 |
| Contains | 包含 |
| NotContains | 不包含 |
| Intersected | 指定值 |
| Empty | 为空 |
| NotEmpty | 不为空 |

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    // 添加一个筛选条件
    await instance.ready();
    const app = instance.Application;
    const filters = await app.Sheets(1).Views(1).Filters;
    const criteria = app.Criteria(1, 'Equals', ['1'])
    const filter = await filters.Add(criteria);

    // 读取添加的筛选条件规则
    const criteria = await app.Sheets(1).Views(1).Filters(1).Criteria
    console.log(criteria.Op) // "Equals"
    console.log(criteria.Field) // 1
    console.log(criteria.Value) // [{type: 'Text', value: '1'}]
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
    // 添加一个筛选条件
    const filters = Application.Sheets(1).Views(1).Filters;
    const criteria = Application.Criteria(1, 'Equals', ['1'])
    const filter = filters.Add(criteria);

    // 读取添加的筛选条件规则
    const criteria = Application.Sheets(1).Views(1).Filters(1).Criteria
    console.log(criteria.Op) // "Equals"
    console.log(criteria.Field) // 1
    console.log(criteria.Value) // [{type: 'Text', value: '1'}]
}
main()
```
 
# 465 API文档 / API / Criteria / Value

本页内容

- 说明
- 浏览器环境示例
- 脚本编辑器示例

# Criteria.Values (属性) ​

## 说明 ​

筛选条件项值

各筛选规则独立地限制了values数组内最多允许填写的元素数，当values内元素数超过阈值时，该筛选规则将失效。“为空、不为空”只允许在op为“Intersected”时填写。“介于”允许最多填写2个元素；“指定值”允许填写65535个元素；其他规则允许最多填写1个元素 values[]数组内的元素为字符串时，表示文本匹配。目前还支持对日期进行动态筛选，此时values[]内的元素需以结构体的形式给出：

```javascript
const dateValue = {"dynamicType": "lastMonth","type": "DynamicSimple"}
Criteria("@日期", "Equals", [dateValue])
```

上述示例对应的筛选条件为“等于上一个月”。 要使用日期动态筛选，values[]内的结构体需要指定"type": "DynamicSimple"，当"op"为"equals"时，"dynamicType"可以为如下的值（大小写不敏感）：

| 枚举值 | 描述 |
| --- | --- |
| today | 今天 |
| yesterday | 昨天 |
| tomorrow | 明天 |
| last7Days | 最近7天 |
| last30Days | 最近30天 |
| thisWeek | 本周 |
| lastWeek | 上周 |
| nextWeek | 下周 |
| thisMonth | 本月 |
| lastMonth | 上月 |
| nextMonth | 次月 |

当"op"为"greater"或"less"时，"dynamicType"只能是昨天、今天或明天。

当"op"为"Intersected"时，value=[{type: 'Any'}]表示全选，value=[{type: '', value: ''}]表示空白

对不同字段类型，values会有不同的用法 联系人字段:

```javascript
// value是一个结构体，指定type为 Contact, value 为用户id
const dateValue = {"type":"Contact", value:"user id"}
```

单/多选项字段:

```javascript
// value是一个结构体，指定type为 SelectItem, value 为选项的ID
const dateValue = {"type":"SelectItem", value:"B"}
```

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    // 添加一个筛选条件
    await instance.ready();
    const app = instance.Application;
    const filters = await app.Sheets(1).Views(1).Filters;
    const criteria = app.Criteria(1, 'Equals', ['1'])
    const filter = await filters.Add(criteria);

    // 读取添加的筛选条件规则
    const criteria = await app.Sheets(1).Views(1).Filters.item(1).Criteria
    console.log(criteria.Op) // "Equals"
    console.log(criteria.Field) // 1
    console.log(criteria.Values) // [{type: 'Text', value: '1'}]
 }
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
    // 添加一个筛选条件
    const filters = Application.Sheets(1).Views(1).Filters;
    const criteria = Application.Criteria(1, 'Equals', ['1'])
    const filter = filters.Add(criteria);

    // 读取添加的筛选条件规则
    const criteria = Application.Sheets(1).Views(1).Filters.item(1).Criteria
    console.log(criteria.Op) // "Equals"
    console.log(criteria.Field) // 1
    console.log(criteria.Values) // [{type: 'Text', value: '1'}]
 }
main()
```
 
# 466 API文档 / API / Filter / Filter对象

本页内容

- 说明
- 方法
- 属性
- 浏览器环境示例
- 脚本编辑器 示例

# Filter (对象) ​

## 说明 ​

单条筛选记录

## 方法 ​

- [Delete](/documents/app-integration-dev/guide/dbsheet/Api/Filter_Delete.html)

## 属性 ​

- [FieldId](/documents/app-integration-dev/guide/dbsheet/Api/Filter_FieldId.html)
- [FilterId](/documents/app-integration-dev/guide/dbsheet/Api/Filter_FilterId.html)
- [Criteria](/documents/app-integration-dev/guide/dbsheet/Api/Filter_Criteria.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const filters = await app.Sheets(1).Views(1).Filters;
    const filter = filters.Item(1);
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const filters = Application.Sheets(1).Views(1).Filters;
    const filter = filters.Item(1);
}
main();
```
 
# 467 API文档 / API / Filter / Criteria

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置筛选条件 ​

Filter.Criteria(属性)

## 说明 ​

可读写

设置或获取 单条筛选记录的条件

![alt text](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/Filter_Criteria.Vsk-X2TT.png)

## 返回值 ​

[Criteria](/documents/app-integration-dev/guide/dbsheet/Api/Criteria.html)

## 浏览器环境示例 ​

javascript

```javascript
// 获取单条筛选记录的条件
async function example() {
    await instance.ready();
    const app = instance.Application;
    // 获取筛选
    const filters = await app.Sheets(1).Views(1).Filters;
    // 获取第一个条件
    const criteria = await filters.Item(1).Criteria;
    console.log(criteria.Op) // "Equals"
    console.log(criteria.Field) // 1
    console.log(criteria.Values[0]) // {type: 'Text', value: '1'}
}

// 设置单条筛选记录的条件
async function example() {
    await instance.ready();
    const app = instance.Application;
    // 构造筛选数据
    const Criteria = app.Criteria(1, "Equals", ["1"])
    // 设置筛选数据
    app.Sheets(1).Views(1).Filters.Item(1).Criteria = Criteria;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const criteria = Application.Sheets(1).Views(1).Filters.Item(1).Criteria;
    criteria = Criteria(1, "Equals", ["1"]);
}
main();
```
 
# 468 API文档 / API / Filter / FieldId

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# Filter.FieldId(属性) ​

## 说明 ​

可读 获取单条筛选记录的字段 ID

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const filters = await app.Sheets(1).Views(1).Filters;
    const fieldId = filters.Item(1).FieldId;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const filters = Application.Sheets(1).Views(1).Filters;
    const fieldId = filters.Item(1).FieldId;
}
main();
```
 
# 469 API文档 / API / Filter / FilterId

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# Filter.FilterId(属性) ​

## 说明 ​

可读 获取单条筛选记录的ID

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const filters = await app.Sheets(1).Views(1).Filters;
    const filterId = filters.Item(1).FilterId;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const filters = Application.Sheets(1).Views(1).Filters;
  const filterId = filters.Item(1).FilterId;
}
main()
```
 
# 470 API文档 / API / Filter / Delete

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 删除筛选条件 ​

Filter.Delete(方法)

## 说明 ​

删除单条筛选记录

## 语法 ​

表达式.Delete()

表达式:Filter

## 参数 ​

无参数

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const res = await app.Sheets(1).Views(1).Filters(1).Delete();
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const res = Application.Sheets(1).Views(1).Filters(1).Delete();
}
main();
```
 
# 471 API文档 / API / Filters / Filters对象

本页内容

- 说明
- 方法
- 属性
- 浏览器环境示例
- 脚本编辑器 示例

# Filters (对象) ​

## 说明 ​

获取指定视图下面的筛选列表

## 方法 ​

- [Item](/documents/app-integration-dev/guide/dbsheet/Api/Filters_Item.html)
- [Add](/documents/app-integration-dev/guide/dbsheet/Api/Filters_Add.html)
- [Clear](/documents/app-integration-dev/guide/dbsheet/Api/Filters_Clear.html)

## 属性 ​

- [Operator](/documents/app-integration-dev/guide/dbsheet/Api/Filters_Operator.html)
- [Count](/documents/app-integration-dev/guide/dbsheet/Api/Filters_Count.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const filters = await app.Sheets(1).Views(1).Filters;
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const filters = Application.Sheets(1).Views(1).Filters;
}
main()
```
 
# 472 API文档 / API / Filters / Count

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# Filters.Count(属性) ​

## 说明 ​

可读 返回筛选列表的个数

## 返回值 ​

Number

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const filters = await app.Sheets(1).Views(1).Filters;
    const count = filters.Count;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const filters = Application.Sheets(1).Views(1).Filters;
  const count = filters.Count;
}
main()
```
 
# 473 API文档 / API / Filters / Operator

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# Filters.Operator(属性) ​

## 说明 ​

可读写

设置或返回头部筛选条件

## 返回值 ​

[FilterOpType](/documents/app-integration-dev/guide/dbsheet/Api/Enum_FilterOpType.html)

## 浏览器环境示例 ​

javascript

```javascript
// 获取筛选的头部筛选条件
async function example() {
    await instance.ready();
    const app = instance.Application;
    const filters = await app.Sheets(1).Views(1).Filters;
    const operator = filters.Operator;
}

// 设置筛选的头部筛选条件
async function example() {
    await instance.ready();
    const app = instance.Application;
    app.Sheets(1).Views(1).Filters.Operator = FilterOpType.Or;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const filters = Application.Sheets(1).Views(1).Filters;
    const operator = filters.Operator;
    operator = FilterOpType.Or;
}
main();
```
 
# 474 API文档 / API / Filters / Add

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 添加筛选条件 ​

Filters.Add(方法)

## 说明 ​

添加筛选条件

![alt text](https://cloudcdn.qwps.cn/open/kingsoft_open_docs/Filter_Criteria.Vsk-X2TT.png)

## 语法 ​

表达式.Add(Criteria)

表达式:Filters

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Criteria | 是 | Object | 筛选条件 |

## 返回值 ​

Filter

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const filters = await app.Sheets(1).Views(1).Filters;
    const criteria = app.Criteria(1, 'Equals', ['1'])
    const filter = await filters.Add(criteria);
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const filters = Application.Sheets(1).Views(1).Filters;
    const criteria = Application.Criteria(1, 'Equals', ['1'])
    const filter = filters.Add(criteria);
}
main()
```
 
# 475 API文档 / API / Filters / Clear

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# Filters.Clear(方法) ​

## 说明 ​

删除视图下面的所有筛选条件或者删除视图下面指定字段类型的所有筛选条件

## 语法 ​

表达式.Clear(FieldId)

表达式:Filters

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| FieldId | 否 | string | 字段ID |

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
// 删除视图下面的所有筛选条件
async function example() {
    await instance.ready();
    const app = instance.Application;
    const filters = await app.Sheets(1).Views(1).Filters;
    const result = filters.Clear();
}

// 同一字段支持多条件筛选时，Clear(fieldId)可以删除该字段下面的所有条件筛选
async function example() {
    await instance.ready();
    const app = instance.Application;
    const filters = await app.Sheets(1).Views(1).Filters;
    const result = filters.Clear('B');
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const filters = Application.Sheets(1).Views(1).Filters;
  const result = filters.Clear();
//   const result = filters.Clear('B');
}
main()
```
 
# 476 API文档 / API / Filters / Item

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# Filters.Item(方法) ​

## 说明 ​

获取指定索引位置的单条筛选条件

## 语法 ​

表达式.Item(Index)

表达式:Filters

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Index | 是 | number | 索引从 1 开始 |

## 返回值 ​

Filter

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const filters = await app.Sheets(1).Views(1).Filters;
    const filter = filters.Item(1);
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const filters = Application.Sheets(1).Views(1).Filters;
    const filter = filters.Item(1);
}
main();
```
 
# 477 API文档 / API / Group / Group对象

本页内容

- 说明
- 方法
- 属性
- 浏览器环境示例
- 脚本编辑器 示例

# Group (对象) ​

## 说明 ​

单条分组记录

## 方法 ​

- [Delete](/documents/app-integration-dev/guide/dbsheet/Api/Group_Delete.html)

## 属性 ​

- [IsAscending](/documents/app-integration-dev/guide/dbsheet/Api/Group_IsAscending.html)
- [Unit](/documents/app-integration-dev/guide/dbsheet/Api/Group_Unit.html)
- [Field](/documents/app-integration-dev/guide/dbsheet/Api/Group_Field.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const res = await app.Sheets(1).Views(1).Groups(1);
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const res = Application.Sheets(1).Views(1).Groups(1);
}
main()
```
 
# 478 API文档 / API / Group / Field

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# Group.Field(属性) ​

## 说明 ​

可读写

分组字段

## 返回值 ​

Field

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const field = await app.Sheets(1).Views(1).Groups(1).Field;

    // 设置分组条件字段
    const group = await app.Sheets(1).Views(1).Groups(1);
    group.Field = '@数字';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const field = Application.Sheets(1).Views(1).Groups(1).Field;

  // 设置分组条件字段
    const group = Application.Sheets(1).Views(1).Groups(1);
    group.Field = '@数字';
}
main()
```
 
# 479 API文档 / API / Group / IsAscending

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# Group.IsAscending(属性) ​

## 说明 ​

可读写

是否为升序

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const isAscending = await app.Sheets(1).Views(1).Groups(1).IsAscending;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const isAscending = Application.Sheets(1).Views(1).Groups(1).IsAscending;
}
main()
```
 
# 480 API文档 / API / Group / Unit

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# Group.Unit(属性) ​

## 说明 ​

可读写

返回单个分组条件单位

只支持设置字段类型为日期/创建时间/最后修改时间的分组单位和值类型为日期类型的公式字段

## 返回值 ​

[DbGroupUnit](/documents/app-integration-dev/guide/dbsheet/Api/Enum_DbGroupUnit.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const unit = await app.Sheets(1).Views(1).Groups(1).Unit;

    // 设置分组单位
    const dateGroup = await app.Sheets(1).Views(1).Groups(3);
    dateGroup.Unit = 'Week';
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const unit = Application.Sheets(1).Views(1).Groups(1).Unit;

     // 设置分组单位
    const dateGroup = Application.Sheets(1).Views(1).Groups(3);
    dateGroup.Unit = 'Week';
}
main();
```
 
# 481 API文档 / API / Group / Delete

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 删除分组 ​

Group.Delete(方法)

## 说明 ​

删除分组条件

## 语法 ​

表达式.Delete()

表达式:Group

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |

## 返回值 ​

ApiResult

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const res = await app.Sheets(1).Views(1).Groups(1).Delete();
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const res = Application.Sheets(1).Views(1).Groups(1).Delete();
}
main()
```
 
# 482 API文档 / API / Groups / Groups对象

本页内容

- 说明
- 方法
- 属性
- 浏览器环境示例
- 脚本编辑器 示例

# Groups (对象) ​

## 说明 ​

获取指定视图下面的分组列表

## 方法 ​

- [Item](/documents/app-integration-dev/guide/dbsheet/Api/Groups_Item.html)
- [Add](/documents/app-integration-dev/guide/dbsheet/Api/Groups_Add.html)
- [FoldAll](/documents/app-integration-dev/guide/dbsheet/Api/Groups_FoldAll.html)
- [UnFoldAll](/documents/app-integration-dev/guide/dbsheet/Api/Groups_UnFoldAll.html)

## 属性 ​

- [IsTempUnFold](/documents/app-integration-dev/guide/dbsheet/Api/Groups_IsTempUnFold.html)
- [Count](/documents/app-integration-dev/guide/dbsheet/Api/Groups_Count.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const groups = await app.Sheets(1).Views(1).Groups;
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const groups = Application.Sheets(1).Views(1).Groups;
}
main()
```
 
# 483 API文档 / API / Groups / Count

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# Groups.Count(属性) ​

## 说明 ​

可读 返回分组列表的个数

## 返回值 ​

Number

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const count = await app.Sheets(1).Views(1).Groups.Count;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const count = await app.Sheets(1).Views(1).Groups.Count;
}
main()
```
 
# 484 API文档 / API / Groups / IsTempUnFold

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# Groups.IsTempUnFold(属性) ​

## 说明 ​

可读 是否支持临时展开分组

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
    const app = instance.Application;
    const isTempUnFold = await app.Sheets(1).Views(1).Groups.IsTempUnFold;
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const isTempUnFold = Application.Sheets(1).Views(1).Groups.IsTempUnFold;
}
main()
```
 
# 485 API文档 / API / Groups / Add

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 添加分组 ​

Groups.Add(方法)

## 说明 ​

添加分组

## 语法 ​

表达式.Add(Field, IsAscending)

表达式:Groups

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Field | 是 | number/string | 新增分组字段索引/新增分组字段ID/新增分组字段名 |
| IsAscending | 否 | boolean | 是否是升序 |

## 返回值 ​

Group

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const res = await app.Sheets(1).Views(1).Groups.Add(1);
    // const res = await app.Sheets(1).Views(1).Groups.Add("@数量");
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const res = Application.Sheets(1).Views(1).Groups.Add(1);
    // const res = Application.Sheets(1).Views(1).Groups.Add("@数量");
}
main();
```
 
# 486 API文档 / API / Groups / FoldAll

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 折叠分组 ​

Groups.FoldAll(方法)

## 说明 ​

折叠分组

## 语法 ​

表达式.FoldAll()

表达式:Groups

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const res = await app.Sheets(1).Views(1).Groups.FoldAll();
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const res = Application.Sheets(1).Views(1).Groups.FoldAll();
}
main()
```
 
# 487 API文档 / API / Groups / Item

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# Groups.Item(方法) ​

## 说明 ​

通过索引位置或者字段ID获取单条分组记录

## 语法 ​

表达式.Item(Index)

表达式:Groups

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Index | 是 | number/string | 获取的分组字段索引/获取的分组字段ID |

## 返回值 ​

Group

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const res = await app.Sheets(1).Views(1).Groups(1);
    // const res = await app.Sheets(1).Views(1).Groups('B');
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const res = Application.Sheets(1).Views(1).Groups(1);
  // const res = Application.Sheets(1).Views(1).Groups('B');
}
main()
```
 
# 488 API文档 / API / Groups / UnFoldAll

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 展开分组 ​

Groups.UnFoldAll(方法)

## 说明 ​

展开分组

## 语法 ​

表达式.UnFoldAll()

表达式:Groups

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const res = await app.Sheets(1).Views(1).Groups.UnFoldAll();
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const res = Application.Sheets(1).Views(1).Groups.UnFoldAll();
}
main()
```
 
# 489 API文档 / API / DbComment / DbComment对象

本页内容

- 说明
- 方法
- 属性
- 浏览器环境示例
- 脚本编辑器 示例

# DbComment (对象) ​

## 说明 ​

DbComment对象的属性和方法可以用来操作该条评论。

## 方法 ​

- [Delete](/documents/app-integration-dev/guide/dbsheet/Api/DbComment_Delete.html)

## 属性 ​

- [TextLinkRuns](/documents/app-integration-dev/guide/dbsheet/Api/TextLinkRuns.html)
- [Id](/documents/app-integration-dev/guide/dbsheet/Api/DbComment_Id.html)
- [Text](/documents/app-integration-dev/guide/dbsheet/Api/DbComment_Text.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const app = instance.Application;
   const recordComment = await app.ActiveView.RecordComments.Item(1)
   recordComment.Add("Hello World")
   const count = recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = await recordComment.Item(i)
    console.log(await comment.Id)
    console.log(await comment.Text)
   }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const recordComment = ActiveView.RecordComments.Item(1)
   recordComment.Add("Hello World")
   const count = recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = recordComment.Item(i)
    console.log(comment.Text)
   }
}
main()
```
 
# 490 API文档 / API / DbComment / Id

本页内容

# RecordComment.Id(属性) ​

## 说明 ​

只读

返回评论的Id

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const app = instance.Application;
   const recordComment = await app.ActiveView.RecordComments.Item(1)
   const count = recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = await recordComment.Item(i)
    console.log(await comment.Id)
   }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const recordComment = ActiveView.RecordComments.Item(1)
   const count = recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = recordComment.Item(i)
    console.log(comment.Id)
   }
}
main()
```
 
# 491 API文档 / API / DbComment / Text

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# RecordComment.Text(属性) ​

## 说明 ​

只读

返回评论的文本

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const app = instance.Application;
   const recordComment = await app.ActiveView.RecordComments.Item(1)
   const count = recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = await recordComment.Item(i)
    console.log(await comment.Text)
   }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const recordComment = ActiveView.RecordComments.Item(1)
   const count = recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = recordComment.Item(i)
    console.log(comment.Text)
   }
}
main()
```
 
# 492 API文档 / API / DbComment / Delete

本页内容

- 说明
- 语法
- 参数
- 返回值
- jsApi 示例
- 脚本编辑器 示例

# DbComment.Delete(方法) ​

## 说明 ​

删除记录

## 语法 ​

表达式.Delete()

表达式:DbComment

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |

## 返回值 ​

boolean

## jsApi 示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const app = instance.Application;
   const recordComment = await app.ActiveView.RecordComments.Item(1)
   const count = recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = await recordComment.Item(i)
    await comment.Delete()
   }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const recordComment = ActiveView.RecordComments(1)
   const count = recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = recordComment.Item(i)
    comment.Delete()
   }
}
main()
```
 
# 493 API文档 / API / RecordComment / RecordComment对象

本页内容

- 说明
- 方法
- 属性
- 浏览器环境示例
- 脚本编辑器 示例

# RecordComment (对象) ​

## 说明 ​

RecordComment对象代表记录Record的评论，可以通过RecordComment对象对记录的评论进行操作。

## 方法 ​

- [Item](/documents/app-integration-dev/guide/dbsheet/Api/RecordComment_Item.html)
- [Add](/documents/app-integration-dev/guide/dbsheet/Api/RecordComment_Add.html)
- [Delete](/documents/app-integration-dev/guide/dbsheet/Api/RecordComment_Delete.html)

## 属性 ​

- [Count](/documents/app-integration-dev/guide/dbsheet/Api/RecordComment_Count.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const app = instance.Application;
   const recordComment = await app.ActiveView.RecordComments.Item(1)
   recordComment.Add("Hello World")
   const count = await recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = await recordComment.Item(i)
    console.log(await comment.Text)
   }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const recordComment = ActiveView.RecordComments.Item(1)
   recordComment.Add("Hello World")
   const count = recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = recordComment.Item(i)
    console.log(comment.Text)
   }
}
main()
```
 
# 494 API文档 / API / RecordComment / Count

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# RecordComment.Count(属性) ​

## 说明 ​

只读 返回的记录里包含评论的数量

## 返回值 ​

Number

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const app = instance.Application;
   const recordComment = await app.ActiveView.RecordComments.Item(1)
   const count = await recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = await recordComment.Item(i)
    console.log(await comment.Text)
   }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const recordComment = ActiveView.RecordComments.Item(1)
   const count = recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = recordComment.Item(i)
    console.log(comment.Text)
   }
}
main()
```
 
# 495 API文档 / API / RecordComment / Add

本页内容

- 说明
- 语法
- 参数
- 返回值
- jsApi 示例
- 脚本编辑器 示例

# 插入新的评论 ​

RecordComment.Add(方法)

## 说明 ​

插入新的评论

## 语法 ​

表达式.Add(Text, TextLinkRuns)

表达式:RecordComment

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Text | 是 | String | 评论的文本，插入到评论的最前方 |
| TextLinkRuns | 否 | Array | 文本的特殊节点属性 |

## 返回值 ​

DbComment

## jsApi 示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const app = instance.Application;
   const recordComment = await app.ActiveView.RecordComments.Item(1)
   recordComment.Add("Hello World")
   const count = await recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = await recordComment.Item(i)
    console.log(await comment.Text)
   }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const recordComment = ActiveView.RecordComments.Item(1)
   recordComment.Add("Hello World")
   const count = recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = recordComment.Item(i)
    console.log(comment.Text)
   }
}
main()
```
 
# 496 API文档 / API / RecordComment / Delete

本页内容

- 说明
- 语法
- 参数
- 返回值
- jsApi 示例
- 脚本编辑器 示例

# 删除评论 ​

RecordComment.Delete(方法)

## 说明 ​

删除评论

## 语法 ​

表达式.Delete()

表达式:RecordComment

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Index | 是 | Number/String | 删除记录的索引 |

## 返回值 ​

ApiResult

## jsApi 示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const app = instance.Application;
   const recordComment = await app.ActiveView.RecordComments.Item(1)
   const count = await recordComment.Count
   for (let i = 1; i <= count; i++) { 
    await recordComment.Delete(i)
   }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const recordComment = ActiveView.RecordComments(1)
   const count = recordComment.Count
   for (let i = 1; i <= count; i++) { 
    await recordComment.Delete(i)
   }
}
main()
```
 
# 497 API文档 / API / RecordComment / Item

本页内容

- 说明
- 语法
- 参数
- 返回值
- jsApi 示例
- 脚本编辑器 示例

# RecordComment.Item(方法) ​

## 说明 ​

获取指定索引位置或评论ID的记录

## 语法 ​

表达式.Item(Index)

表达式:RecordComment

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Index | 否 | number/string | 传入number时索引从1开始，传入字符串时表示评论id |

## 返回值 ​

[DbComment](/documents/app-integration-dev/guide/dbsheet/Api/DbComment.html)

## jsApi 示例 ​

javascript

```javascript
async function example() {
   await instance.ready();
   const app = instance.Application;
   const recordComment = await app.ActiveView.RecordComments(1)
   const count = await recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = await recordComment.Item(i)
    console.log(await comment.Text)
   }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
   const recordComment = ActiveView.RecordComments(1)
   const count = recordComment.Count
   for (let i = 1; i <= count; i++) { 
    const comment = recordComment.Item(i)
    console.log(comment.Text)
   }
}
main()```
```
 
# 498 API文档 / API / RecordComments / RecordComments对象

本页内容

- 说明
- 方法
- 属性
- 事件
- 浏览器环境示例
- 脚本编辑器 示例

# RecordComments (对象) ​

## 说明 ​

RecordComments 当前视图上的评论，每条记录对应一个RecordComments对象，RecordComments对象的属性和方法可以用来操作该条记录的评论。 多维表的评论是在记录上的，评论的对象主体上是三级 RecordComments -&gt; RecordComment -&gt; DbComment

## 方法 ​

- [Item](/documents/app-integration-dev/guide/dbsheet/Api/RecordComments_Item.html)

## 属性 ​

- [Count](/documents/app-integration-dev/guide/dbsheet/Api/RecordComments_Count.html)

## 事件 ​

- [OnCreate](/documents/app-integration-dev/guide/dbsheet/Api/RecordComments_OnCreate.html)
- [OnDelete](/documents/app-integration-dev/guide/dbsheet/Api/RecordComments_OnDelete.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const comment = await app.ActiveView.RecordComments(1).Item(5)
 console.log(await comment.Text) // 输出 (3) ['aaaa', 45470, 1]
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
 const comment = Application.ActiveView.RecordComments(1).Item(5)
 console.log(comment.Text) // 输出 (3) ['aaaa', 45470, 1]
}
main()
```
 
# 499 API文档 / API / RecordComments / Count

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# RecordComments.Count(属性) ​

## 说明 ​

可读写

返回当前RecordComment的数量，因为每条记录都会有一个RecordComment对象，所以返回的是记录的数量。

## 返回值 ​

Number

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const count = await instance.Application.ActiveView.RecordComments.Count
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const count = Application.ActiveView.RecordComments.Count
}
main()
```
 
# 500 API文档 / API / RecordComments / Item

本页内容

- 说明
- 语法
- 参数
- 返回值
- jsApi 示例
- 脚本编辑器 示例

# RecordComments.Item(方法) ​

## 说明 ​

通过记录的索引或ID获取指定记录的评论对象。可以简化写法 RecordComments.Item(1) 可以简化为 RecordComments(1)

## 语法 ​

表达式.Item(Index)

表达式:RecordComments

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Index | 否 | number/string | 传入number时记录索引从1开始，传入字符串时表示记录id |

## 返回值 ​

Self

## jsApi 示例 ​

javascript

```javascript
async function example() {
  await instance.ready();
  const app = instance.Application;
  const count = await app.ActiveView.RecordComments.Item(1).Count
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const values = Application.ActiveView.RecordComments.Item(1).Count
}
main()```
```
 
# 501 API文档 / API / RecordComments / OnCreate

本页内容

- 说明
- 语法
- 参数
- 返回值
- 事件返回数据
- 浏览器环境示例

# 监听插入评论 ​

RecordComments.OnCreate(方法)

## 说明 ​

为 RecordComments 添加 OnCreate 事件,当创建 评论 时触发。注意在脚本编辑器中使用时，脚本运行结束就会退出运行，这时可能回调无法被正常触发 这个方法只能监听视图的事件， 如果在浏览器环境需要全局监听也可以使用

javascript

```javascript
jssdk.on("OnBroadcast", async (res) => {
    const data = res.Data
    if (data.type == "DB_COMMENT_UPDATE") { // 收到文档评论更新消息
        if (data.shouldNotLocalUpdate) {
            // 本地更新评论信息
            console.log("收到广播消息：", data)
            const info = data.info
            const {sheetStId, commentId, recordId, action} = info
            if (action == "Add") {
                // 新增评论
                const addText = await jssdk.Application.Sheets.ItemById(sheetStId).ActiveView.RecordComments(recordId).Item(commentId).Text
                console.log("新增评论：", addText)
            } else if (action == "Delete") {
                // 删除评论
                console.log("删除评论：", info)
            }
        }
    }
})
```

可以通过 action 来判断是哪个事件触发的

## 语法 ​

表达式.OnDelete(Callback)

表达式: RecordComments

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Callback | 是 | func | 指定事件发生时的回调方法 ; const eventContext = await RecordComments.OnCreate(()=&gt;{ ... }) |

## 返回值 ​

EventContext

## 事件返回数据 ​

DbComment

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    let eventContext;
    eventContext = await app.ActiveView.RecordComments.OnCreate(async (data)=> {
        const info = data.info
        const {sheetStId, commentId, recordId, action} = info
        if (action == "Add") {
            // 新增评论
            const addText = await app.Sheets.ItemById(sheetStId).ActiveView.RecordComments(recordId).Item(commentId).Text
            console.log("新增评论：", addText)
        } else if (action == "Delete") {
            // 删除评论
            console.log("删除评论：", info)
        }
    })

    // 移除监听
    // eventContext.Destroy();
}
```
 
# 502 API文档 / API / RecordComments / OnDelete

本页内容

- 说明
- 语法
- 参数
- 返回值
- 事件返回数据
- 事件返回数据示例
- 浏览器环境示例

# 监听删除评论 ​

RecordComments.OnDelete(方法)

## 说明 ​

为 RecordComments 添加 Delete 事件,当删除 评论 时触发。注意在脚本编辑器中使用时，脚本运行结束就会退出运行，这时可能回调无法被正常触发 这个方法只能监听视图的事件， 如果在浏览器环境需要全局监听也可以使用

jssdk.on("OnBroadcast", (res)=&gt;console.error("##", res))

回调的消息数据包含的内容跟事件返回数据是一致的， 可以通过 action 来判断是哪个事件触发的

## 语法 ​

表达式.OnDelete(Callback)

表达式: RecordComments

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Callback | 是 | func | 指定事件发生时的回调方法 ; const eventContext = await RecordComments.OnDelete(()=&gt;{ ... }) |

## 返回值 ​

EventContext

## 事件返回数据 ​

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| commentId | String | 评论ID |
| recordId | String | 记录ID |
| sheetStId | Number | 表ID |

## 事件返回数据示例 ​

```javascript
{"recordId":"Bk","sheetStId":1,"commentId":"e66e42020baa4d5455da5d2043c631a5","action":"Delete"}
```

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    let eventContext;
    eventContext = await app.ActiveView.RecordComments.OnDelete((data)=>console.error(JSON.stringify(data)))
    await app.ActiveView.RecordComments(1).Item(1).Delete()

    // 移除监听
    // eventContext.Destroy();
}
```
 
# 503 API文档 / API / DataSource / DataSource对象

本页内容

- 说明
- 方法
- 浏览器环境示例
- 脚本编辑器 示例

# DataSource (对象) ​

## 说明 ​

DataSource 对象

## 方法 ​

- [CreateSyncDBSheets](/documents/app-integration-dev/guide/dbsheet/Api/DataSource_CreateSyncDBSheets.html)
- [ImportFromCloud](/documents/app-integration-dev/guide/dbsheet/Api/DataSource_ImportFromCloud.html)
- [ImportFromLocal](/documents/app-integration-dev/guide/dbsheet/Api/DataSource_ImportFromLocal.html)
- [CreateSummarySheet](/documents/app-integration-dev/guide/dbsheet/Api/DataSource_CreateSummarySheet.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const dataSource = await app.DataSource;
    const sheets = await dataSource.ImportFromCloud('100127684526')
    console.log(sheets)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const dataSource = Application.DataSource
    const sheets = await dataSource.ImportFromCloud('100127684526')
    console.log(sheets)
}
main()
```
 
# 504 API文档 / API / DataSource / CreateSummarySheet

本页内容

- 说明
- 语法
- 参数
- 返回值
- jsAPI示例
- 脚本编辑器 示例

# 创建合并表 ​

DataSource.CreateSummarySheet(方法)

JSSDK v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

将数据源表合并成一个合并表。

## 语法 ​

表达式.CreateSummarySheet(SummarySourceConfigs)

表达式:DataSource

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| SummarySourceConfigs | 是 | SummarySourceConfigs | 数据源配置对象 |

## 返回值 ​

SummarySheet

## jsAPI示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    // 通过SummarySourceConfigs对象构造符合规范的数据源配置对象，来创建合并表
    const configs = await app.SummarySourceConfigs;
    await configs.Add('100138603654', [1])
    const sheet = await app.DataSource.CreateSummarySheet(configs)
    console.log(sheet)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  // 通过SummarySourceConfigs对象构造符合规范的数据源配置对象，来创建合并表
  const configs = Application.SummarySourceConfigs;
  configs.Add('100138603654', [1])
  const sheet = Application.DataSource.CreateSummarySheet(configs)
  console.log(sheet)
}
main()
```
 
# 505 API文档 / API / DataSource / CreateSyncDBSheets

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 创建同步表 ​

DataSource.CreateSyncDBSheets(方法)

## 说明 ​

创建同步表

## 语法 ​

表达式.CreateSyncDBSheets(FileId, OfficeType, SheetIds) 表达式:DataSource

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| FileId | 是 | string | 在线表格文件的 fileId |
| OfficeType | 是 | 'd' | 'k' | 在线表格文件的类型：d 代表多维表格，k 代表智能表格 |
| SheetIds | 是 | number[] | 在线表格文件的 stId 数组 |

## 返回值 ​

SyncDBSheet[]

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const syncSheets = await app.DataSource.CreateSyncDBSheets('100127684526', 'd', [1])
    console.log(syncSheets)
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  const syncSheets = Application.DataSource.CreateSyncDBSheets('100127684526', 'd', [1])
  console.log(syncSheets)
}
main()
```
 
# 506 API文档 / API / DataSource / ImportFromCloud

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# DataSource.ImportFromCloud(方法) ​

## 说明 ​

将在线表格文件导入为新的数据表

## 语法 ​

表达式.ImportFromCloud(FileId)

表达式:DataSource

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| FileId | 是 | string | 在线表格文件的 fileId |

## 返回值 ​

DbSheet[]

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    try {
        const sheets = await app.DataSource.ImportFromCloud('100127684526')
        console.log(sheets)
    } catch (error) {
        console.log('导入数据表失败', error.message)
    }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    try {
        const sheets = Application.DataSource.ImportFromCloud('100127684526')
        console.log(sheets)
    } catch (error) {
        console.log('导入数据表失败', error.message)
    }
}
main()
```
 
# 507 API文档 / API / DataSource / ImportFromLocal

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# DataSource.ImportFromLocal(方法) ​

## 说明 ​

将本地表格文件导入为新的数据表

## 语法 ​

表达式.ImportFromLocal(File)

表达式:DataSource

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| File | 是 | File | 本地表格文件对象 |

## 返回值 ​

DbSheet[]

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];
    if (!file) {
        console.error('没有选择文件');
        return;
    }
    
    try {
        const sheets = await app.DataSource.ImportFromLocal(file)
        console.log(sheets)
    } catch (error) {
        console.log('导入数据表失败', error.message)
    }
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
  try {
    const sheets = Application.DataSource.ImportFromLocal(file)
    console.log(sheets)
  } catch (error) {
    console.log('导入数据表失败', error.message)
  }
}
main()
```
 
# 508 API文档 / API / SummarySheet / SummarySheet对象

本页内容

- 说明
- 方法
- 属性
- jsAPI示例
- 脚本编辑器示例

# SummarySheet (对象) ​

JSSDK v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

SummarySheet 合并表对象，用于操作合并表

## 方法 ​

- [RefreshSyncSheet](/documents/app-integration-dev/guide/dbsheet/Api/SummarySheet_RefreshSyncSheet.html)
- [RemoveSheetSyncLink](/documents/app-integration-dev/guide/dbsheet/Api/SummarySheet_RemoveSheetSyncLink.html)

## 属性 ​

- [SourceConfigs](/documents/app-integration-dev/guide/dbsheet/Api/SummarySheet_SourceConfigs.html)

## jsAPI示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 // 先切到合并表，然后获取合并表实例
 const summarySheet = await app.ActiveSheet
 const configs = await summarySheet.SourceConfigs
 console.log(await configs.Item(1).FileId)
 }
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 // 先切到合并表，然后获取合并表
 const summarySheet = Application.ActiveSheet
 const configs = summarySheet.SourceConfigs
 console.log(configs.Item(1).FileId)
}
main()
```
 
# 509 API文档 / API / SummarySheet / SourceConfigs

本页内容

- 说明
- 返回值
- jsAPI示例
- 脚本编辑器示例

# SummarySheet.SourceConfigs(属性) ​

JSSDK v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

只读 合并表的数据源配置对象

## 返回值 ​

SummarySourceConfigs

## jsAPI示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 // 先切到合并表，然后获取合并表配置
 const configs = await app.ActiveSheet.SourceConfigs
 console.log(await configs.Item(1).FileId)
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 // 先切到合并表，然后获取合并表配置
 const configs = Application.ActiveSheet.SourceConfigs
 console.log(configs.Item(1).FileId)
}
main()
```
 
# 510 API文档 / API / SummarySheet / RefreshSyncSheet

本页内容

# SummarySheet.RefreshSyncSheet(方法) ​

JSSDK v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

刷新数据

## 语法 ​

表达式.RefreshSyncSheet()

表达式:SummarySheet

## 参数 ​

无参数

## 返回值 ​

Boolean

## jsAPI示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 // 切到某个合并表，获取该合并表的实例对象
 const summarySheet = await app.ActiveSheet
 await summarySheet.RefreshSyncSheet()
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 // 切到某个合并表，获取该合并表的实例对象
 const summarySheet = Application.ActiveSheet
 summarySheet.RefreshSyncSheet()
}
main()
```
 
# 511 API文档 / API / SummarySheet / RemoveSheetSyncLink

本页内容

- 说明
- 语法
- 参数
- 返回值
- jsAPI示例
- 脚本编辑器示例

# SummarySheet.RemoveSheetSyncLink(方法) ​

JSSDK v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

解除同步关系

## 语法 ​

表达式.RemoveSheetSyncLink()

表达式:SummarySheet

## 参数 ​

无参数

## 返回值 ​

Boolean

## jsAPI示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 // 切到某个合并表，获取该合并表的实例对象
 const summarySheet = await app.ActiveSheet
 await summarySheet.RemoveSheetSyncLink()
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 // 切到某个合并表，获取该合并表的实例对象
 const summarySheet = Application.ActiveSheet
 summarySheet.RemoveSheetSyncLink()
}
main()
```
 
# 512 API文档 / API / SummarySourceConfig / SummarySourceConfig对象

本页内容

- 说明
- 方法
- 属性
- jsAPI示例
- 脚本编辑器示例

# SummarySourceConfig (对象) ​

JSSDK v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

合并表数据源对象中的某个源文件配置对象，数据源对象是由多个文件配置组成的一个数组。 每个源文件配置对象都包含了两个属性，文件id：FileId，文件中选中的数据表id数组：SheetIds

## 方法 ​

- [SetUrl](/documents/app-integration-dev/guide/dbsheet/Api/SummarySourceConfig_SetUrl.html)
- [SetSheets](/documents/app-integration-dev/guide/dbsheet/Api/SummarySourceConfig_SetSheets.html)

## 属性 ​

- [FileId](/documents/app-integration-dev/guide/dbsheet/Api/SummarySourceConfig_FileId.html)
- [SheetIds](/documents/app-integration-dev/guide/dbsheet/Api/SummarySourceConfig_SheetIds.html)

## jsAPI示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 // 先切到合并表，然后获取合并表配置
 const config = await app.ActiveSheet.SourceConfigs.Item(1)
 // 返回配置文件的id
 console.log(await config.FileId)
 }
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 // 先切到合并表，然后获取合并表配置
 const config = Application.ActiveSheet.SourceConfigs.Item(1)
 // 返回配置文件的id
 console.log(config.FileId)
}
main()
```
 
# 513 API文档 / API / SummarySourceConfig / FileId

本页内容

- 说明
- 返回值
- jsAPI示例
- 脚本编辑器示例

# SummarySourceConfig.FileId(属性) ​

JSSDK v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

可读写 合并表中某个源文件配置对象的文件id。注意：修改后，该对象中对应的SheetIds会自动重置为空数组。（只是编辑的本地对象，需要调SummarySourceConfigs对象中的Apply方法才能更新到云上）

## 返回值 ​

String

## jsAPI示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const configs = await app.ActiveSheet.SourceConfigs
 const config = await configs.Item(1)
 console.log(await config.FileId)
 config.FileId = '100139340929'
 console.log(await config.FileId)
 config.SheetIds = [1, 3]
 // 将改动后的配置更新到云上
 await configs.Apply()
 }
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 const configs = Application.ActiveSheet.SourceConfigs
 const config = configs.Item(1)
 console.log(config.FileId)
 config.FileId = '100139340929'
 console.log(config.FileId)
 config.SheetIds = [1, 3]
 // 将改动后的配置更新到云上
 configs.Apply()
}
main()
```
 
# 514 API文档 / API / SummarySourceConfig / SheetIds

本页内容

- 说明
- 返回值
- jsAPI示例
- 脚本编辑器示例

# SummarySourceConfig.SheetIds(属性) ​

JSSDK v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

可读写 合并表中某个源文件配置对象的选中数据表的id数组。（只是编辑的本地对象，需要调SummarySourceConfigs对象中的Apply方法才能更新到云上）

## 返回值 ​

Number[]

## jsAPI示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const configs = await app.ActiveSheet.SourceConfigs
 const config = await configs.Item(1)
 console.log(await config.SheetIds)
 config.SheetIds = [1, 3]
 console.log(await config.SheetIds)
 // 将改动后的配置更新到云上
 await configs.Apply()
 }
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 const configs = Application.ActiveSheet.SourceConfigs
 const config = configs.Item(1)
 console.log(config.SheetIds)
 config.SheetIds = [1, 3]
 console.log(config.SheetIds)
 // 将改动后的配置更新到云上
 configs.Apply()
}
main()
```
 
# 515 API文档 / API / SummarySourceConfig / SetSheets

本页内容

- 说明
- 语法
- 参数
- 返回值
- jsAPI示例
- 脚本编辑器示例

# SummarySourceConfig.SetSheets(方法) ​

JSSDK v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

通过一组数据表名或索引，来修改合并表源文件配置对象中的SheetIds属性（只是编辑的本地对象，需要调SummarySourceConfigs对象中的Apply方法才能更新到云上）

## 语法 ​

表达式.SetSheets(Sheets)

表达式:SummarySourceConfig

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Sheets | 是 | Number/String[] | 数据表名或索引组成的数组 |

## 返回值 ​

Boolean

## jsAPI示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const configs = await app.ActiveSheet.SourceConfigs
 const config = await configs.Item(1)
 console.log(await config.SheetIds)
 await config.SetSheets([1, '数据表(2)'])
 console.log(await config.SheetIds)
 // 将改动后的配置更新到云上
 await configs.Apply()
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 const configs = Application.ActiveSheet.SourceConfigs
 const config = configs.Item(1)
 console.log(config.SheetIds)
 config.SetSheets([1, '数据表(2)'])
 console.log(config.SheetIds)
 // 将改动后的配置更新到云上
 configs.Apply()
}
main()
```
 
# 516 API文档 / API / SummarySourceConfig / SetUrl

本页内容

- 说明
- 语法
- 参数
- 返回值
- jsAPI示例
- 脚本编辑器示例

# SummarySourceConfig.SetUrl(方法) ​

JSSDK v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

通过url修改合并表源文件配置对象中的FileId属性（只是编辑的本地对象，需要调SummarySourceConfigs对象中的Apply方法才能更新到云上）

## 语法 ​

表达式.SetUrl(Url)

表达式:SummarySourceConfig

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Url | 是 | String | 文件url |

## 返回值 ​

Boolean

## jsAPI示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const configs = await app.ActiveSheet.SourceConfigs
 const config = await configs.Item(1)
 console.log(await config.FileId)
 await config.SetUrl("https://www.kdocs.cn/l/ccSvj0b4dC7u?R=L1MvMTI=")
 console.log(await config.FileId)
 await config.SetSheets([1, 2])
 // 将改动后的配置更新到云上
 await configs.Apply()
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 const configs = Application.ActiveSheet.SourceConfigs
 const config = configs.Item(1)
 console.log(config.FileId)
 config.SetUrl("https://www.kdocs.cn/l/ccSvj0b4dC7u?R=L1MvMTI=")
 console.log(config.FileId)
 config.SetSheets([1, 2])
 // 将改动后的配置更新到云上
 configs.Apply()
}
main()
```
 
# 517 API文档 / API / SummarySourceConfigs / SummarySourceConfigs对象

本页内容

- 说明
- 方法
- 属性
- jsAPI示例
- 脚本编辑器示例

# SummarySourceConfigs (对象) ​

JSSDK v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

合并表的数据源配置对象

## 方法 ​

- [Item](/documents/app-integration-dev/guide/dbsheet/Api/SummarySourceConfigs_Item.html)
- [Add](/documents/app-integration-dev/guide/dbsheet/Api/SummarySourceConfigs_Add.html)
- [Delete](/documents/app-integration-dev/guide/dbsheet/Api/SummarySourceConfigs_Delete.html)
- [Apply](/documents/app-integration-dev/guide/dbsheet/Api/SummarySourceConfigs_Apply.html)

## 属性 ​

- [Count](/documents/app-integration-dev/guide/dbsheet/Api/SummarySourceConfigs_Count.html)

## jsAPI示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 // 先切到合并表，然后获取合并表配置
 const configs = await app.ActiveSheet.SourceConfigs
 console.log(await configs.Item(1).FileId)
 }
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 // 先切到合并表，然后获取合并表配置
 const configs = Application.SourceConfigs
 console.log(configs.Item(1).FileId)
}
main()
```
 
# 518 API文档 / API / SummarySourceConfigs / Count

本页内容

- 说明
- 返回值
- jsAPI示例
- 脚本编辑器示例

# SummarySourceConfigs.Count(属性) ​

JSSDK v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

只读 合并表数据源配置中，包含的源文件的数量

## 返回值 ​

Number

## jsAPI示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const configs = await app.ActiveSheet.SourceConfigs
 console.log(await configs.Count)
 // 向配置中新加一个源文件配置
 await configs.Add("100136699885", [1, 2])
 console.log(await configs.Count)
 }
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 const configs = app.ActiveSheet.SourceConfigs
 console.log(configs.Count)
 // 向配置中新加一个源文件配置
 configs.Add("100136699885", [1, 2])
 console.log(configs.Count)
}
main()
```
 
# 519 API文档 / API / SummarySourceConfigs / Add

本页内容

- 说明
- 语法
- 参数
- 返回值
- jsAPI示例
- 脚本编辑器示例

# 添加数据源 ​

SummarySourceConfigs.Add(方法)

JSSDK v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

合并表配置对象中添加数据源（只是编辑的本地对象，需要调Apply方法才能更新到云上）

## 语法 ​

表达式.Add(File,Sheets)

表达式:SummarySourceConfigs

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| File | 是 | String | 文件id/文件url |
| Sheets | 是 | Number/String[] | 数据表数组，支持两种形式，表名和索引（从1开始） |

## 返回值 ​

SummarySourceConfig

## jsAPI示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const configs = await app.ActiveSheet.SourceConfigs
 const config = await configs.Add("100136699885", [1, 2])
 // 返回配置文件的id
 console.log(await config.FileId)
 // 返回文件下选中的数据表id数组
 console.log(await config.SheetIds)
 // 将改动后的配置更新到云上
 await configs.Apply()
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 const configs = Application.ActiveSheet.SourceConfigs
 const config = configs.Add("100136699885", [1, 2])
 // 返回配置文件的id
 console.log(config.FileId)
 // 返回文件下选中的数据表id数组
 console.log(config.SheetIds)
 // 将改动后的配置更新到云上
 configs.Apply()
}
main()
```
 
# 520 API文档 / API / SummarySourceConfigs / Apply

本页内容

- 说明
- 语法
- 参数
- 返回值
- jsAPI示例
- 脚本编辑器示例

# SummarySourceConfigs.Apply(方法) ​

JSSDK v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

更新本地的数据源配置到云上

## 语法 ​

表达式.Apply()

表达式:SummarySourceConfigs

## 参数 ​

无参数

## 返回值 ​

Boolean

## jsAPI示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const configs = await app.ActiveSheet.SourceConfigs
 await configs.Add("100136699885", [1, 2])
 await configs.Apply()
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 const configs = Application.ActiveSheet.SourceConfigs
 configs.Add("100136699885", [1, 2])
 configs.Apply()
}
main()
```
 
# 521 API文档 / API / SummarySourceConfigs / Delete

本页内容

- 说明
- 语法
- 参数
- 返回值
- jsAPI示例
- 脚本编辑器示例

# 删除数据源 ​

SummarySourceConfigs.Delete(方法)

JSSDK v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

合并表配置对象中删除数据源（只是编辑的本地对象，需要调Apply方法才能更新到云上）

## 语法 ​

表达式.Delete(Index)

表达式:SummarySourceConfigs

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Index | 是 | Number/String | 支持索引、文件id和文件url |

## 返回值 ​

Boolean

## jsAPI示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const configs = await app.ActiveSheet.SourceConfigs
 await configs.Delete(1)
 // 将改动后的配置更新到云上
 await configs.Apply()
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 const configs = Application.ActiveSheet.SourceConfigs
 configs.Delete(1)
 // 将改动后的配置更新到云上
 configs.Apply()
}
main()
```
 
# 522 API文档 / API / SummarySourceConfigs / Item

本页内容

- 说明
- 语法
- 参数
- 返回值
- jsAPI示例
- 脚本编辑器示例

# SummarySourceConfigs.Item(方法) ​

JSSDK v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

获取指定索引的源文件配置

## 语法 ​

表达式.Item(Index)

表达式:SummarySourceConfigs

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| Index | 是 | number/string | 索引从1开始/文件id |

## 返回值 ​

SummarySourceConfig

## jsAPI示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const config = await app.ActiveSheet.SourceConfigs.Item(1)
 // 返回配置文件的id
 console.log(await config.FileId)
 // 返回文件下选中的数据表id数组
 console.log(await config.SheetIds)
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 const config = Application.ActiveSheet.SourceConfigs.Item(1)
 // 返回配置文件的id
 console.log(config.FileId)
 // 返回文件下选中的数据表id数组
 console.log(config.SheetIds)
}
main()
```
 
# 523 API文档 / API / SyncDBSheet / SyncDBSheet对象

本页内容

- 说明
- 方法
- jsAPI示例
- 脚本编辑器示例

# SyncDBSheet (对象) ​

JSSDK v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

所有同步表的基类，包含同步表共有的操作：刷新同步表和解除同步关系

## 方法 ​

- [RefreshSyncSheet](/documents/app-integration-dev/guide/dbsheet/Api/SyncDBSheet_RefreshSyncSheet.html)
- [RemoveSheetSyncLink](/documents/app-integration-dev/guide/dbsheet/Api/SyncDBSheet_RemoveSheetSyncLink.html)

## jsAPI示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 // 切到某个同步表，获取该同步表的实例对象
 const syncSheet = await app.ActiveSheet
 await syncSheet.RefreshSyncSheet()
 }
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 // 切到某个同步表，获取该同步表的实例对象
 const syncSheet = Application.ActiveSheet
 syncSheet.RefreshSyncSheet()
}
main()
```
 
# 524 API文档 / API / SyncDBSheet / RefreshSyncSheet

本页内容

- 说明
- 语法
- 参数
- 返回值
- jsAPI示例
- 脚本编辑器示例

# 刷新数据 ​

SyncDBSheet.RefreshSyncSheet(方法)

JSSDK v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

刷新数据

## 语法 ​

表达式.RefreshSyncSheet()

表达式:SyncDBSheet

## 参数 ​

无参数

## 返回值 ​

Boolean

## jsAPI示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 // 切到某个同步表，获取该同步表的实例对象
 const syncSheet = await app.ActiveSheet
 await syncSheet.RefreshSyncSheet()
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 // 切到某个同步表，获取该同步表的实例对象
 const syncSheet = Application.ActiveSheet
 syncSheet.RefreshSyncSheet()
}
main()
```
 
# 525 API文档 / API / SyncDBSheet / RemoveSheetSyncLink

本页内容

- 说明
- 语法
- 参数
- 返回值
- jsAPI示例
- 脚本编辑器示例

# 解除同步关系 ​

SyncDBSheet.RemoveSheetSyncLink(方法)

JSSDK v1.1.10+、WebOffice v2.4.1+ 支持

## 说明 ​

解除同步关系

## 语法 ​

表达式.RemoveSheetSyncLink()

表达式:SyncDBSheet

## 参数 ​

无参数

## 返回值 ​

Boolean

## jsAPI示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 // 切到某个同步表，获取该同步表的实例对象
 const syncSheet = await app.ActiveSheet
 await syncSheet.RemoveSheetSyncLink()
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 // 切到某个同步表，获取该同步表的实例对象
 const syncSheet = Application.ActiveSheet
 syncSheet.RemoveSheetSyncLink()
}
main()
```
 
# 526 API文档 / API / NoticeBar / NoticeBar对象

本页内容

- 说明
- 方法
- 属性
- 浏览器环境示例

# NoticeBar (对象) ​

## 说明 ​

公告栏对象

## 方法 ​

## 属性 ​

- [Visible](/documents/app-integration-dev/guide/dbsheet/Api/NoticeBar_Visible.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const noticeBar = await app.Window.NoticeBar
 }
```
 
# 527 API文档 / API / NoticeBar / Visible

本页内容

- 说明
- 返回值
- 浏览器环境示例

# 展开/收起 ​

NoticeBar.Visible(属性)

## 说明 ​

可读写

设置公告栏可见性

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function getVisible() {
    await instance.ready();
    const app = instance.Application;
    const noticeBar = await app.Window.NoticeBar
    const visible =  await noticeBar.Visible     // 为true公告栏隐藏，为false公告栏显示
}
async function setVisible() {
    await instance.ready();
    const app = instance.Application;
    const noticeBar = await app.Window.NoticeBar
    noticeBar.Visible = false     // 隐藏公告栏
    noticeBar.Visible = true      // 显示公告栏
}
```
 
# 528 API文档 / API / ApiResult / ApiResult对象

本页内容

- 说明
- 方法
- 属性
- 浏览器环境示例
- 脚本编辑器示例

# ApiResult (对象) ​

## 说明 ​

API调用后的返回值

## 方法 ​

## 属性 ​

- [Code](/documents/app-integration-dev/guide/dbsheet/Api/ApiResult_Code.html)
- [Message](/documents/app-integration-dev/guide/dbsheet/Api/ApiResult_Message.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const field = await app.Sheets(1).Views(1).Fields(1)
 const result = await field.Move(3) 
 console.log(result.Code)
 console.log(result.Message)
 }
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
 const field = Application.Sheets(1).Views(1).Fields(1)
 const result = field.Move(3) 
 console.log(result.Code)
 console.log(result.Message)
 }
main()
```
 
# 529 API文档 / API / ApiResult / Code

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器示例

# API调用返回执行结果Code ​

ApiResult.Code(属性)

## 说明 ​

只读 API调用返回执行结果

## 返回值 ​

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const field = await app.Sheets(1).Views(1).Fields(1)
    const result = await field.Move(3) 
    if(result.Code !== 0) {
        console.log(result.Message)
    } 
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
    const field = Application.Sheets(1).Views(1).Fields(1)
    const result = field.Move(3) 
    if(result.Code !== 0) {
        console.log(result.Message)
    }
}
main()
```
 
# 530 API文档 / API / ApiResult / Message

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器示例

# API调用返回执行结果Message ​

ApiResult.Message(属性)

## 说明 ​

只读 API调用返回执行结果

## 返回值 ​

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const field = await app.Sheets(1).Views(1).Fields(1)
    const result = await field.Move(3) 
    if(result.Code !== 0) {
        console.log(result.Message)
    } 
}
```

## 脚本编辑器示例 ​

javascript

```javascript
function main() {
    const field = Application.Sheets(1).Views(1).Fields(1)
    const result = field.Move(3) 
    if(result.Code !== 0) {
        console.log(result.Message)
    }
}
main()
```

```javascript

```
 
# 531 API文档 / API / 枚举值 / DbAutomationPresetType

本页内容

- DbAutomationPresetType

## DbAutomationPresetType ​

| 枚举值 | 字符串值 | 描述 |
| --- | --- | --- |
| DbAutomationPresetType.CheckedNotifyContact | "CheckedNotifyContact" | 状态打勾时通知联系人 |
| DbAutomationPresetType.UpdatedNotifyContact | "UpdatedNotifyContact" | 内容变更时通知联系人 |
| DbAutomationPresetType.DueDateNotifyContact | "DueDateNotifyContact" | 到期时通知联系人 |
 
# 532 API文档 / API / 枚举值 / DbButtonIcon

本页内容

- DbButtonIcon

## DbButtonIcon ​

按钮字段的图标枚举值

| 枚举值 | 字符串值 | 描述 |
| --- | --- | --- |
| DbButtonIcon.GestureThumb | 'gesture / thumb' |  |
| DbButtonIcon.PaperPlane | 'paper / plane' |  |
| DbButtonIcon.Insert | 'insert' |  |
| DbButtonIcon.Star | 'star' |  |
| DbButtonIcon.Stamp | 'stamp' |  |
| DbButtonIcon.Camera | 'camera' |  |
| DbButtonIcon.Bubble | 'bubble' |  |
| DbButtonIcon.BubbleTwo | 'bubble / two' |  |
| DbButtonIcon.Bell | 'bell' |  |
| DbButtonIcon.Bulb | 'bulb' |  |
| DbButtonIcon.Heart | 'heart' |  |
| DbButtonIcon.Gift | 'gift' |  |
| DbButtonIcon.PeopleDouble | 'people / double' |  |
| DbButtonIcon.Info | 'info' |  |
| DbButtonIcon.Successful | 'successful' |  |
| DbButtonIcon.Stick | 'stick' |  |
| DbButtonIcon.Cell | 'cell' |  |
| DbButtonIcon.CalendarCheckIn | 'calendar / check / in' |  |
| DbButtonIcon.Print | 'print' |  |
| DbButtonIcon.None | 'none' |  |
 
# 533 API文档 / API / 枚举值 / DbFieldValueType

本页内容

- DbFieldValueType

## DbFieldValueType ​

公式字段的数字格式

| 枚举值 | 字符串值 | 描述 |
| --- | --- | --- |
| DbFieldValueType.Fvt / Text | 'Fvt / Text' | 文本 |
| DbFieldValueType.Fvt / Number | 'Fvt / Number' | 数字 |
| DbFieldValueType.Fvt / Contact | "Fvt / Contact" | 联系人 |
| DbFieldValueType.Fvt / Date | "Fvt / Date" | 日期 |
| DbFieldValueType.Fvt / Time | "Fvt / Time" | 时间 |
| DbFieldValueType.Fvt / Logic | "Fvt / Logic" | 逻辑 |
 
# 534 API文档 / API / 枚举值 / DbFilterCriteriaOpType

本页内容

- DbFilterCriteriaOpType

## DbFilterCriteriaOpType ​

筛选操作的枚举值

| 枚举值 | 字符串值 | 描述 | 适用字段 |
| --- | --- | --- | --- |
| DbFilterCriteriaOpType.Null | "Null" | 无 | 所有 |
| DbFilterCriteriaOpType.Equals | "Equals" | 等于 | 数字，日期，文本，集合 |
| DbFilterCriteriaOpType.NotEqu | "NotEqu" | 不等于 | 数字，日期，文本，集合 |
| DbFilterCriteriaOpType.Greater | "Greater" | 大于 | 数字，日期 |
| DbFilterCriteriaOpType.GreaterEqu | "GreaterEqu" | 大于等于 | 数字，日期 |
| DbFilterCriteriaOpType.Less | "Less" | 小于 | 数字，日期 |
| DbFilterCriteriaOpType.LessEqu | "LessEqu" | 小于等于 | 数字，日期 |
| DbFilterCriteriaOpType.GreaterEquAndLessEqu | "GreaterEquAndLessEqu" | 介于 | 数字，日期 |
| DbFilterCriteriaOpType.LessOrGreater | "LessOrGreater" | 不介于 | 数字，日期 |
| DbFilterCriteriaOpType.BeginWith | "BeginWith" | 开始是 | 文本 |
| DbFilterCriteriaOpType.NotBeginWith | "NotBeginWith" | 开始不是 | 文本 |
| DbFilterCriteriaOpType.EndWith | "EndWith" | 结束是 | 文本 |
| DbFilterCriteriaOpType.NotEndWith | "NotEndWith" | 结束不是 | 文本 |
| DbFilterCriteriaOpType.Contains | "Contains" | 包含 | 文本 |
| DbFilterCriteriaOpType.NotContains | "NotContains" | 不包含 | 文本 |
| DbFilterCriteriaOpType.Intersected | "Intersected" | 有交集 | 集合 |
| DbFilterCriteriaOpType.NotIntersected | "NotIntersected" | 没有交集 | 集合 |
| DbFilterCriteriaOpType.SubsetOf | "SubsetOf" | 子集 | 集合 |
| DbFilterCriteriaOpType.SupersetOf | "SupersetOf" | 超集 | 集合 |
| DbFilterCriteriaOpType.Empty | "Empty" | 空 | 数字，日期，文本，集合 |
| DbFilterCriteriaOpType.NotEmpty | "NotEmpty" | 非空 | 数字，日期，文本，集合 |
 
# 535 API文档 / API / 枚举值 / DbGroupUnit

本页内容

- DbGroupUnit

## DbGroupUnit ​

| 枚举值 | 字符串值 | 描述 |
| --- | --- | --- |
| DbGroupUnit.Text | 'Text' | 文本 |
| DbGroupUnit.Year | 'Year' | 年 |
| DbGroupUnit.Month | 'Month' | 月 |
| DbGroupUnit.Week | 'Week' | 周 |
| DbGroupUnit.Day | 'Day' | 天 |
| DbGroupUnit.Hour | 'Hour' | 小时 |
| DbGroupUnit.Minute | 'Minute' | 分钟 |
| DbGroupUnit.Second | 'Second' | 秒 |
 
# 536 API文档 / API / 枚举值 / DbLookupFunction

本页内容

- DbLookupFunction

## DbLookupFunction ​

统计引用的方法

| 枚举值 | 字符串值 | 描述 |
| --- | --- | --- |
| DbLookupFunction.Origin | 'Origin' | - |
| DbLookupFunction.Sum | 'Sum' | 求和 |
| DbLookupFunction.Counta | 'Counta' | 计数 |
| DbLookupFunction.Average | 'Average' | 平均值 |
| DbLookupFunction.Max | 'Max' | 最大值 |
| DbLookupFunction.Min | 'Min' | 最小值 |
| DbLookupFunction.Unique | 'Unique' | 去重 |
| DbLookupFunction.CountaUnique | 'CountaUnique' | 去重计数 |
| DbLookupFunction.ToString | 'ToString' | 连接字符串 |
 
# 537 API文档 / API / 枚举值 / DbSharedCriteriaType

本页内容

- DbSharedCriteriaType

## DbSharedCriteriaType ​

| 枚举值 | 字符串值 | 描述 |
| --- | --- | --- |
| DbSharedCriteriaType.Custom | "Custom" | 自定义 |
| DbSharedCriteriaType.All | "All" | 所有 |
| DbSharedCriteriaType.None | "None" | 无 |
| DbSharedCriteriaType.Self | "Self" | 仅自己 |
 
# 538 API文档 / API / 枚举值 / FilterOpType

本页内容

- FilterOpType

## FilterOpType ​

| 枚举值 | 字符串值 | 描述 |
| --- | --- | --- |
| FilterOpType.Or | "Or" | 满足任一条件 |
| FilterOpType.And | "And" | 满足所有条件 |
 
# 539 API文档 / API / 枚举值 / SharedLinkPermissionType

本页内容

- SharedLinkPermissionType

## SharedLinkPermissionType ​

| 枚举值 | 字符串值 |  | 描述 |
| --- | --- | --- | --- |
| SharedLinkPermissionType.edit | "edit" | 可编辑 |  |
| SharedLinkPermissionType.read | "read" | 可查看 |  |
| SharedLinkPermissionType.addAndDel | "addAndDel" | 可添加,删除 |  |
 
# 540 API文档 / API / 枚举值 / SharedLinkToType

本页内容

- SharedLinkToType

## SharedLinkToType ​

| 枚举值 | 字符串值 |  | 描述 |
| --- | --- | --- | --- |
| SharedLinkToType.assigned | "assigned" | 指定人 |  |
| SharedLinkToType.anyone | "anyone" | 所有人 |  |
| SharedLinkToType.company | "company" | 企业内成员 |  |
| SharedLinkToType.onlylinkcreator | "onlylinkcreator" | 链接创建者 |  |
 
# 541 API文档 / API / Font / Font对象

本页内容

# Font (对象) ​

## 说明 ​

字体对象

## 属性 ​

- [Color](/documents/app-integration-dev/guide/dbsheet/Api/Font_Color.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const font = await app.Sheets(1).Views(1).RecordRange(10).Font
    console.log(await font.Color)
    font.Color = "#00ff00"
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const font = Application.Sheets(1).Views(1).RecordRange(10).Font
    console.log(font.Color)
    font.Color = "#00ff00"
 }
main()
```
 
# 542 API文档 / API / Font / Color

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 字体颜色属性 ​

Font.Color(属性)

## 说明 ​

可读写

设置或者获取字体颜色属性

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const font = await app.Sheets(1).Views(1).RecordRange(10).Font
    console.log(await font.Color)
    font.Color = "#00ff00"
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const font = Application.Sheets(1).Views(1).RecordRange(10).Font
    console.log(font.Color)
    font.Color = "#00ff00"
}
main();
```
 
# 543 API文档 / API / Interior / Interior对象

本页内容

- 说明
- 属性
- 浏览器环境示例
- 脚本编辑器 示例

# Interior (对象) ​

## 说明 ​

单元格内部填充对象

## 属性 ​

- [Color](/documents/app-integration-dev/guide/dbsheet/Api/Interior_Color.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const interior = await app.Sheets(1).Views(1).RecordRange(10).Interior
    console.log(await interior.Color)
    interior.Color = "#00ff00"
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const interior = Application.Sheets(1).Views(1).RecordRange(10).Interior
    console.log(interior.Color)
    interior.Color = "#00ff00"
 }
main()
```
 
# 544 API文档 / API / Interior / Color

本页内容

- 说明
- 返回值
- 浏览器环境示例
- 脚本编辑器 示例

# 设置背景颜色 ​

Interior.Color(属性)

## 说明 ​

可读写

设置背景颜色属性

## 返回值 ​

String

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const interior = await app.Sheets(1).Views(1).RecordRange(10).Interior
    console.log(await interior.Color)
    interior.Color = "#00ff00"
}
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const interior = Application.Sheets(1).Views(1).RecordRange(10).Interior
    console.log(interior.Color)
    interior.Color = "#00ff00"
}
main();
```
 
# 545 API文档 / API / Navigator / Navigator对象

本页内容

- 说明
- 方法
- 属性
- 浏览器环境示例

# Navigator (对象) ​

## 说明 ​

导航栏对象

## 方法 ​

## 属性 ​

- [Visible](/documents/app-integration-dev/guide/dbsheet/Api/Navigator_Visible.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const navigator = await app.Window.Navigator
 }
```
 
# 546 API文档 / API / Navigator / Visible

本页内容

- 说明
- 返回值
- 浏览器环境示例

# 设置导航栏可见性 ​

Navigator.Visible(属性)

## 说明 ​

可读写

设置导航栏隐藏或显示

## 返回值 ​

Boolean

## 浏览器环境示例 ​

javascript

```javascript
async function getVisible() {
    await instance.ready();
    const app = instance.Application;
    const navigator = await app.Window.Navigator
    const visible =  await navigator.Visible     // 为true导航栏隐藏，为false导航栏显示
}
async function setVisible() {
    await instance.ready();
    const app = instance.Application;
    const navigator = await app.Window.Navigator
    navigator.Visible = false     // 隐藏导航栏
    navigator.Visible = true      // 显示导航栏
}
```
 
# 547 API文档 / API / Window / Window对象

本页内容

- 说明
- 方法
- 属性
- 浏览器环境示例
- 脚本编辑器 示例

# Window (对象) ​

## 说明 ​

窗口对象，窗口对象只能在浏览器内使用，在脚本编辑器由于是在服务器上运行的脚本，不适用于Window对象

## 方法 ​

- [SetLayout](/documents/app-integration-dev/guide/dbsheet/Api/Window_SetLayout.html)
- [BailHook](/documents/app-integration-dev/guide/dbsheet/Api/Window_BailHook.html)
- [DisplayRecord](/documents/app-integration-dev/guide/dbsheet/Api/Window_DisplayRecord.html)
- [HiddenAllRecord](/documents/app-integration-dev/guide/dbsheet/Api/Window_HiddenAllRecord.html)

## 属性 ​

- [Navigator](/documents/app-integration-dev/guide/dbsheet/Api/Window_Navigator.html)
- [NoticeBar](/documents/app-integration-dev/guide/dbsheet/Api/Window_NoticeBar.html)
- [GanttViewUI](/documents/app-integration-dev/guide/dbsheet/Api/Window_GanttViewUI.html)

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
 await instance.ready();
 const app = instance.Application;
 const windowObject = await app.Window
 }
```

## 脚本编辑器 示例 ​

javascript

```javascript
function main() {
    const windowObject = Window
}
main()
```
 
# 548 API文档 / API / Window / GanttViewUI

本页内容

- 说明
- 返回值
- 浏览器环境示例

# 甘特视图的界面设置 ​

Window.GanttViewUI(属性)

## 说明 ​

返回甘特视图的界面设置，如果当前视图不是甘特视图，则返回undefined

## 返回值 ​

GanttViewUI

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const navigator = await app.Window.GanttViewUI
}
```
 
# 549 API文档 / API / Window / Navigator

本页内容

- 说明
- 返回值
- 浏览器环境示例

# 窗口导航栏 ​

Window.Navigator(属性)

## 说明 ​

返回窗口导航栏对象信息

## 返回值 ​

Navigator

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const navigator = await app.Window.Navigator
}
```
 
# 550 API文档 / API / Window / NoticeBar

本页内容

- 说明
- 返回值
- 浏览器环境示例

# 窗口公告栏 ​

Window.NoticeBar(属性)

## 说明 ​

返回窗口公告栏对象信息

## 返回值 ​

NoticeBar

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    const noticeBar = await app.Window.NoticeBar
}
```
 
# 551 API文档 / API / Window / BailHook

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例

# 代理界面元素 ​

Window.BailHook(方法)

## 说明 ​

对特定的界面元素进行代理，代理方法返回true时，则不显示原界面

## 语法 ​

表达式.BailHook(CmbId)

表达式: Window

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| CmbId | 是 | string | 界面元素ID，目前支持 `RecordInfo`(记录详情卡片) |

## 返回值 ​

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;

    const hook = await app.Window.BailHook("RecordInfo")
    hook.InvokeSingle((params)=>{
        console.log(params) // 移动端返回参数 {recordId: 'Jp', activeFieldId: 'E'}
                            // pc端返回参数 {recordId: 'Jp', isShowComment: false}
                            // 移动端和PC端 用到的参数都是 recordId，其它参数是界面区别
        // 可以通过 await WPSOpenApi.Application.ActiveView.RecordRange(params.recordId).Value 读取记录的数据
        const count = await Application.ActiveView.RecordRange.Count
        const record = Application.ActiveView.RecordRange(params.recordId)
        const values = await record.Value
        const index = await record.Index // 注意 Index base 1, 可能有多条记录，返回 [index]
        const prevRecord = Application.ActiveView.RecordRange(index.map(_=> _ - 1))
        const nextRecord = Application.ActiveView.RecordRange(index.map(_=> _ + 1))
        // 在这里实现自定义的界面逻辑，替换掉原来的界面
        record.Select()
        return true // return true 会不弹出原界面
    })
}
```
 
# 552 API文档 / API / Window / DisplayRecord

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例

# 展开记录 ​

## 说明 ​

当前窗口下展开记录，显示详情信息

## 语法 ​

表达式.DisplayRecord(RecodId)

表达式: Window

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| RecodId | 否 | String | 展开的记录ID |

## 返回值 ​

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    await app.Window.DisplayRecord("B");
}
```
 
# 553 API文档 / API / Window / HiddenAllRecord

本页内容

- 说明
- 语法
- 返回值
- 浏览器环境示例

# 关闭当前展开的记录 ​

## 说明 ​

关闭当前展开的记录

## 语法 ​

表达式.HiddenAllRecord()

表达式: Window

## 返回值 ​

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    await app.Window.HiddenAllRecord();
}
```
 
# 554 API文档 / API / Window / SetLayout

本页内容

- 说明
- 语法
- 参数
- 返回值
- 浏览器环境示例

# 设置经典布局 ​

Window.SetLayout(方法)

## 说明 ​

设置该视图是否为经典布局

## 语法 ​

表达式.SetLayout(isClassic)

表达式: Window

## 参数 ​

| 参数名 | 是否必需 | 类型 | 描述 |
| --- | --- | --- | --- |
| isClassic | 是 | Boolean | 是否为经典布局 |

## 返回值 ​

## 浏览器环境示例 ​

javascript

```javascript
async function example() {
    await instance.ready();
    const app = instance.Application;
    await app.Window.SetLayout(true);
    await app.Window.SetLayout(false);
}
```
