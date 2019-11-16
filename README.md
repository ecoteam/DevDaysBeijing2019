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

1. 打开目标文件夹，在命令提示符中，输入以下命令创建项目。

    ```cmd
    yo office react ReactAddInDemo excel  --ts
    ```

2. 输入以下命令以在Visual Studio Code中打开项目。

    ```cmd
    cd ReactAddInDemo
    code .
    ```

   您也可以启动Visual Studio Code后，执行以下步骤来打开项目。
   * 选择File -> Open Folder
   * 选择ReactAddInDemo目录


3. 打开 `src\taskpane\components\App.tsx` 文件，将
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
             ["Orange", 98],
             ["Banana", 109],
             ["Peach", 173],
             ["Grapefruit", 182],
             ["Strawberry", 60],
         ]);
       });
     } catch (error) {
       console.error(error);
     }
   };
   ```
4. 在Visual Studio Code中按下 `Ctrl+` ` 键，打开命令提示符，执行以下命令以启动Excel并加载Add-in。

   ```cmd
   npm start
   ```

5. 在Add-in中按下 `Run` 按钮，可以看到，在工作簿中插入了一些股票数据。

6. 在 `src\taskpane\components\` 路径下创建 `Data.tsx` 文件，将其替换为以下代码。
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
           let sheet = context.workbook.worksheets.getItem("Sample");
           if(this.props.rangeAddress != null && this.props.rangeAddress != ""){
               let range = sheet.getRange(this.props.rangeAddress);
               let chart = sheet.charts.add(Excel.ChartType.pie, range, Excel.ChartSeriesBy.columns);
               chart.dataLabels.showValue = false;
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
               iconProps={
                 {iconName: "Chart"}
                }
               onClick={this.click}
             >
               Add Chart
             </Button>
         </section>
       );
     }
   }
   ```

7. 打开 `src\taskpane\components\App.tsx` ，将其替换为以下代码。
   ```typescript
   import * as React from "react";
   import { Button, ButtonType } from "office-ui-fabric-react";
   import Header from "./Header";
   import HeroList, { HeroListItem } from "./HeroList";
   import Progress from "./Progress";
   import Data from "./Data";
 
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
           sheet.name = "Sample";
           let table = sheet.tables.add('A1:B1', true);
           table.getHeaderRowRange().values = [["Company", "Price"]];
           table.rows.add(null, [
             ["Rrange", 98],
             ["Banana",190],
             ["Peach", 173],
             ["Grapefruit", 182],
             ["Strawberry", 60]
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
               iconProps={
                 {iconName: "ChevronRight"}
                }
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

8. 如果Excel已经被关闭，在Visual Studio Code中按下 `Ctrl+` ` 键，打开命令提示符，执行以下命令以启动Excel并加载Add-in。

   ```cmd
   npm start
   ```

   如果Excel仍然开着，刷新Add-in。

9. 在Add-in中依次按下 `>Add Data` 按钮和 `>Add Chart` 按钮，可以看到，在工作簿中插入了一些股票数据，以及相对应的图表。

# 使用Custom Functions

在本节中，您将打开ScriptLab创建一个Custom Function，并通过其从Web Service中获取数据。

1. 选择 Insert -> Get Add-ins 打开 Office Add-ins 窗口。在搜索窗口输入 Script 后搜索并安装 Script Lab。

2. 在Ribbon中找到Script Lab，点击 Code 打开任务面板。

3. 在Sample中搜索Custom，找到Custom Function的Sample，选择Basic Custom Function。

4. 将其名称改为 DevDays。这将作为Custom Function的命令空间。

5. 点击Register注册。

6. 在Excel单元格中输入`=SCRIPTLAB.DEVDAYS.SPHEREVOLUME(2)`后回车。可以看到其进行了计算并返回了半径为2的球的体积。

7. 回到Code任务面板，将DevDays中的代码替换为以下代码。

    ```typescript
    /**
    * Get stock price.
    * @customfunction getStockPrice
    * @param symbol Stock symbol.
    * @returns Stock price.
    */
    export function getStockPrice(symbol: string): Promise<string> {
      let url = "https://mockstockprice.azurewebsites.net/stocks/" + symbol;
      return new Promise((resolve, reject) => {
        fetch(url)
          .then(function (response) {
            if (response.status != 200) {
              reject();
              return;
            }

            return response.json();
          })
          .then(function (json) {
            resolve(json.price);
          })
      })
    }
    ```

8. 在Excel的B1单元格中，输入`Contoso.getStockPrice(A1)`。

9. 在A1单元格中输入`MSFT`，可以看到B1单元格自动进行了重算，并且获取了微软的股票数据。

10. 回到Code任务面板，将DevDays中的代码替换为以下代码。

    ```typescript
    /**
    * Get stock price.
    * @customfunction getStockPrice
    * @param symbol Stock symbol.
    * @returns Stock price.
    * @streaming
    */
    export function getStockPrice(symbol: string, invocation: CustomFunctions.StreamingInvocation<string>) {
      let url = "https://mockstockprice.azurewebsites.net/stocks/" + symbol;
      const timer = setInterval(() => {
        fetch(url)
          .then(function (response) {
            if (response.status != 200) {
              throw new Error();
            }

            return response.json();
          })
          .then(function (json) {
            invocation.setResult(json.price);
          })
      }, 3000);

      invocation.onCanceled = () => {
        clearInterval(timer);
      }
    }
    ```

11. 打开之前创建的基于React的Add-in，插入数据和图表。

12. 在B2单元格中，输入`Contoso.getStockPrice(A2)`。

13. 将B2单元格填充至B3:B6。可以看到Custom Functions和Excel内置函数一样，可以自动更新参数，并自动进行计算。

13. Excel中之前输入的信息现在将会自动刷新。图表也会随着数据的刷新而变化。