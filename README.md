本课程将向您展示如何使用React构建Office Add-in，及如何使用Custom Functions。

# 环境配置

您需要安装以下程序:
* [Node.js](https://nodejs.org/) (V8.0.0+)
* [Visual Studio Code](https://code.visualstudio.com/download)

您需要安装以下NPM Packages:
* [Yeoman](https://github.com/yeoman/yo)
* [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office)

要安装NPM Packages，请启动命令提示符并运行以下命令:

```cmd
npm install -g yo generator-office
```

# 使用React构建Office Add-in

在本节中，您将使用Yeoman generator for Office Add-ins创建一个基于React的Office Add-in项目。您还将使用Excel JavaScript API创建一个表格，并基于它创建图表。

1. 在命令提示符中，输入以下命令创建项目。

    ```cmd
    yo office
    ```

2. 根据提示和下方信息进行选择。

* Choose a project type: `Office Add-in Task Pane project using React framework`
* Choose a script type: `TypeScript`
* What do you want to name your add-in? `DevDaysBeijing2019`
* Which Office client application would you like to support? `Excel`

3. 输入以下命令以在Visual Studio Code中打开项目。

    ```cmd
    cd DevDaysBeijing2019
    code .
    ```

   您也可以启动Visual Studio Code后，执行以下步骤来打开项目。
   * 选择File -> Open Folder
   * 选择DevDaysBeijing2019目录


4. 打开 `src\components\App.tsx` 文件，将
  ```typescript
  click = async () => {
    try {
      await Excel.run(async context => {
        /**
         * Insert your Excel code here
         */
        const range = context.workbook.getSelectedRange();

        // Read the range address
        range.load("address");

        // Update the fill color
        range.format.fill.color = "yellow";

        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  };
  ```
  替换为
  ```typescript
  click = async () => {
    try {
      await Excel.run(async context => {
        let sheet = context.workbook.worksheets.getFirst();
        let expensesTable = sheet.tables.add('A1:B1', true);
        expensesTable.getHeaderRowRange().values = [["Company","Price"]];
        expensesTable.rows.add(null, [
            ["orange", 98],
            ["Banana", 109],
            ["Peach", 173],
            ["Grapefruit", 182],
            ["strawberry", 60],
        ]);
        let range = expensesTable.getRange();
        range.load("Address");
        await context.sync();

        this.setState({rangeAddress: range.address});
      });
    } catch (error) {
      console.error(error);
    }
  };
  ```
5. 在Visual Studio Code中按下 `Ctrl+` ` 键，打开命令提示符，执行以下命令以启动Excel并加载Add-in。

  ```cmd
  npm start
  ```
6. 在Add-in中按下 `Run` 按钮，可以看到，在工作簿中插入了一些股票数据。

7. 打开 `src\components\App.tsx` ，将其替换为以下代码。
  ```typescript
    import * as React from "react";
    import { Button, ButtonType } from "office-ui-fabric-react";
    import Header from "./Header";
    import HeroList, { HeroListItem } from "./HeroList";
    import Progress from "./Progress";
    import Data from "./Data";
    /* global Button, console, Excel, Header, HeroList, HeroListItem, Progress */

    export interface AppProps {
      title: string;
      isOfficeInitialized: boolean;
    }

    export interface AppState {
      listItems: HeroListItem[];
      rangeAddress: string;
    }

    export default class App extends React.Component<AppProps, AppState> {
      constructor(props, context) {
        super(props, context);
        this.state = {
          listItems: [],
          rangeAddress: ""
        };
      }

      componentDidMount() {
        this.setState({
          listItems: [
            {
              icon: "Ribbon",
              primaryText: "Achieve more with Office integration"
            },
            {
              icon: "Unlock",
              primaryText: "Unlock features and functionality"
            },
            {
              icon: "Design",
              primaryText: "Create and visualize like a pro"
            }
          ]
        });
      }

      click = async () => {
        try {
          await Excel.run(async context => {
            let sheet = context.workbook.worksheets.add();
            context.workbook.worksheets.getFirst().delete();
            sheet.activate();
            let table = sheet.tables.add('A1:B1', true);
            table.getHeaderRowRange().values = [["Company", "Price"]];
            table.rows.add(null, [
              ["orange", 98],
              ["banana",190],
              ["peach", 173],
              ["grapefruit", 182],
              ["strawberry", 60]
            ]);
            let range = table.getRange();
            range.load("Address");
            await context.sync();

            this.setState({rangeAddress: range.address});
          });
        } catch (error) {
          console.error(error);
        }
      };

      render() {
        const { title, isOfficeInitialized } = this.props;

        if (!isOfficeInitialized) {
          return (
            <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
          );
        }

        return (
          <div className="ms-welcome">
            <Header logo="assets/logo-filled.png" title={this.props.title} message="Welcome" />
            <HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}>
              <p className="ms-font-l">
                Modify the source files, then click <b>Add Data</b>.
              </p>
              <Button
                className="ms-welcome__action"
                buttonType={ButtonType.hero}
                iconProps={{ iconName: "ChevronRight" }}
                onClick={this.click}
              >
                Add Data
              </Button>
            </HeroList>
            <Data rangeAddress={this.state.rangeAddress}></Data>
          </div>
        );
      }
    }
  ```

7. 在 `src\components\` 路径下创建 `Data.tsx` 文件，将其替换为以下代码。
  ```typescript
    import * as React from "react";
    import { Button, ButtonType } from "office-ui-fabric-react";

    export interface DataProps {
      rangeAddress: string;
    }

    export default class Data extends React.Component<DataProps> {
      click = async () => {    
        //const { rangeAddress } = this.props;
        try {
          await Excel.run(async context => {
            /**
            * Insert your Excel code here
            */
            let sheet = context.workbook.worksheets.getLast();
            if(this.props.rangeAddress != null && this.props.rangeAddress != ""){
                let range = sheet.getRange(this.props.rangeAddress);
                let chart = sheet.charts.add(Excel.ChartType.line, range, Excel.ChartSeriesBy.columns);
                chart.dataLabels.showValue = true;
            }
          });
        } catch (error) {
          console.error(error);
        }
      };

      render() {
        return (
          <section className="ms-welcome__main">
            <Button
                className="ms-welcome__features"
                buttonType={ButtonType.hero}
                iconProps={{ iconName: "Chart" }}
                onClick={this.click}
              >
                Add Chart
              </Button>
          </section>
        );
      }
    }
  ```

8. 如果Excel已经被关闭，在Visual Studio Code中按下 `Ctrl+` ` 键，打开命令提示符，执行以下命令以启动Excel并加载Add-in。

    ```cmd
    npm start
    ```

   如果Excel仍然开着，刷新Add-in。

9. 在Add-in中依次按下 `>Add Data` 按钮和 `>Add Chart` 按钮，可以看到，在工作簿中插入了一些股票数据，以及相对应的图表。

# 集成Custom Functions

在本节中，您将在上一节建立的Office Add-in项目基础上，使用Custom Functions从Web Service中获取数据。

1. 打开`manifest.xml`文件，将其替换为以下代码。

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>9f760d38-beba-4707-8b36-10af2ba1a820</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="My Office Add-in"/>
  <Description DefaultValue="A template to get started."/>
  <IconUrl DefaultValue="https://localhost:3000/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="https://localhost:3000/assets/icon-80.png"/>
  <SupportUrl DefaultValue="https://www.contoso.com/help"/>
  <AppDomains>
    <AppDomain>contoso.com</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="CustomFunctionsRuntime" MinVersion="1.1"/>
    </Sets>
  </Requirements>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:3000/taskpane.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <AllFormFactors>
          <ExtensionPoint xsi:type="CustomFunctions">
            <Script>
              <SourceLocation resid="Functions.Script.Url"/>
            </Script>
            <Page>
              <SourceLocation resid="Functions.Page.Url"/>
            </Page>
            <Metadata>
              <SourceLocation resid="Functions.Metadata.Url"/>
            </Metadata>
            <Namespace resid="Functions.Namespace"/>
          </ExtensionPoint>
        </AllFormFactors>
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="CommandsGroup">
                <Label resid="CommandsGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="TaskpaneButton">
                  <Label resid="TaskpaneButton.Label"/>
                  <Supertip>
                    <Title resid="TaskpaneButton.Label"/>
                    <Description resid="TaskpaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonId1</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
        <bt:Url id="Functions.Metadata.Url" DefaultValue="https://localhost:3000/dist/functions.json"/>
        <bt:Url id="Functions.Page.Url" DefaultValue="https://localhost:3000/dist/functions.html"/>
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="https://go.microsoft.com/fwlink/?LinkId=276812"/>
        <bt:Url id="Commands.Url" DefaultValue="https://localhost:3000/commands.html"/>
        <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Functions.Namespace" DefaultValue="CONTOSO"/>
        <bt:String id="GetStarted.Title" DefaultValue="Get started with your sample add-in!"/>
        <bt:String id="CommandsGroup.Label" DefaultValue="Commands Group"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="Show Taskpane"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="Your sample add-in loaded succesfully. Go to the HOME tab and click the 'Show Taskpane' button to get started."/>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Click to Show a Taskpane"/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

2. 在Visual Studio Code中按下 `Ctrl+` ` 键，打开命令提示符，执行以下命令。

```cmd
npm install -S custom-functions-metadata-plugin @types/custom-functions-runtime
```

3. 在`src`目录下，根据以下目录结构创建`functions.html`和`functions.ts`文件。

```
src
 ├── commands
 ├── functions
 │    ├── functions.html
 │    └── functions.ts
 └── taskpane
```

4. 打开`functions.html`文件，粘贴以下代码。

```html
<!-- Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License. -->

<!DOCTYPE html>
<html>

<head>
  <meta charset="UTF-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
  <meta http-equiv="Expires" content="0" />
  <title></title>
  <script src="https://appsforoffice.microsoft.com/lib/1.1/hosted/custom-functions-runtime.js" type="text/javascript"></script>
</head>

<body>
    
</body>

</html>
```

5. 打开`functions.ts`文件，粘贴以下代码。

```typescript
/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
/* global clearInterval, console, setInterval */

export function add(first: number, second: number): number {
  return first + second;
}

/**
 * Displays the current time once a second.
 * @customfunction
 * @param invocation Custom function handler
 */
export function clock(invocation: CustomFunctions.StreamingInvocation<string>): void {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time.
 * @returns String with the current time formatted for the current locale.
 */
export function currentTime(): string {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param incrementBy Amount to increment
 * @param invocation Custom function handler
 */
export function increment(incrementBy: number, invocation: CustomFunctions.StreamingInvocation<number>): void {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param message String to write.
 * @returns String to write.
 */
export function logMessage(message: string): string {
  console.log(message);

  return message;
}
```

6. 在Visual Studio Code中按下 `Ctrl+` ` 键，打开命令提示符，执行以下命令以启动Excel并加载Add-in。

```cmd
npm start
```

7. 在Excel中，在任意单元格中输入`=Contoso.Add(1,2)`。此时您将看到其进行了计算，并且显示了计算结果`3`。

8. 打开`functions.ts`文件，在文件最后粘贴以下代码。

```typescript
```