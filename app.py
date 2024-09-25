from flask import Flask, render_template, request, jsonify
import os
import traceback
from model import Fdata,renderexcel
from flask_cors import CORS
import webview
import threading

app = Flask(__name__)
CORS(app)

# EXCEL_FILE = 'form_data.xlsx'
latest_file = "空.xlsx"
sheetmap = {}
xspreadsheetmap = {}

@app.route('/')
def index():
    return render_template('index.html', tables=20)

def start_server():
    app.run(port=5000)

@app.route('/submit', methods=['POST'])
def submit():
    global latest_file
    global xspreadsheetmap
    global sheetmap
    try:
        data = request.form.to_dict()
        # print(data)
        # data = {'projectname': '11', 'projectplace': '11', 'projectind': '11', 'input1': '8', 'input2': '2', 'productionLoad3': '60', 'productionLoad4': '70', 'productionLoad5': '80', 'productionLoad6': '90', 'productionLoad7': '100', 'productionLoad8': '100', 'benchmarkyield': '10', 'benchmark_s_paybackperiod': '6', 'benchmark_d_paybackperiod': '6', 'install_rate': '10', 'trans_rate': '0', 'pre_rate': '10', 'preup_rate': '10', 'additionalCostName_1': '11', 'additionalCostAmount_1': '1000', 'additionalCostPrice_1': '10000', 'additionalCostEquipment_1': '1000', 'additionalCostName_2': '22', 'additionalCostAmount_2': '1000', 'additionalCostPrice_2': '10000', 'additionalCostEquipment_2': '1000', 'additionalotherCostName_1': '33', 'additionalOtherCost_1': '1000', 'additionalOtherCostType_1': 'tudi', 'additionalotherCostName_2': '44', 'additionalOtherCost_2': '800', 'additionalOtherCostType_2': 'intangible', 'additionalotherCostName_3': '55', 'additionalOtherCost_3': '900', 'additionalOtherCostType_3': 'other', 'days1': '30', 'days2': '20', 'days3': '20', 'days4': '20', 'days5': '30', 'days6': '30', 'days7': '20', 'days8': '5', 'days9': '70', 'loadrate': '4', 'repayMethod': 'ave_cap_int', 'payTime': '5', 'investAmount_1': '4005', 'investPercentage_1': '50', 'loanAmount_1': '2403', 'loanPercentage_1': '60', 'dep_year1': '10', 'res_rate1': '10', 'dep_year2': '10', 'res_rate2': '10', 'dep_year3': '10', 'depreciationMethod': 'DDB', 'amo_year1': '5', 'amo_year2': '5', 'production1': '100', 'production2': '60', 'production3': '100', 'production4': '3', 'production5': '100', 'production6': '10', 'production7': '10', 'production8': '4', 'production9': '100', 'production10': '140', 'production11': '17', 'production12': '4', 'production13': '2', 'production14': '25', 'production15': '10', 'production16': '10', 'projectcostCounter': '2', 'othercostCounter': '3', 'constInvenstCounter': '1', 'additionalCosts': '{"11":"1000","22":"1000","33":"1000","44":"800","55":"900"}', 'investmentPlan': '{"1":{"amount":"4005","percentage":"50"}}', 'LoanPlan': '{"1":{"amount":"2403","percentage":"60"},"2":{"amount":"2403","percentage":"60"}}', 'leftInvestment': '4005', 'leftInvestmentPercentage': '50.0'}

        f = Fdata(data)
        latest_file, sheetmap,xspreadsheetmap = f.toCalExcel() #写入excel
        # data_secure = {'projectname': '示例', 'projectplace': '示例', 'projectind': '示例', 'input1': '6', 'input2': '2',
        #                'productionLoad3': '60', 'productionLoad4': '80', 'productionLoad5': '100',
        #                'productionLoad6': '100', 'benchmarkyield': '1', 'benchmark_s_paybackperiod': '4',
        #                'benchmark_d_paybackperiod': '4', 'install_rate': '10', 'trans_rate': '10', 'pre_rate': '10',
        #                'preup_rate': '10', 'additionalCostName_1': '主要项目', 'additionalCostAmount_1': '1000',
        #                'additionalCostPrice_1': '100000', 'additionalCostEquipment_1': '10000',
        #                'additionalCostName_2': '辅助项目', 'additionalCostAmount_2': '1000',
        #                'additionalCostPrice_2': '100000', 'additionalCostEquipment_2': '10000',
        #                'additionalotherCostName_1': '土地征用费', 'additionalOtherCost_1': '100000',
        #                'additionalOtherCostType_1': '', 'days1': '20', 'days2': '20', 'days3': '20', 'days4': '20',
        #                'days5': '20', 'days6': '20', 'days7': '20', 'days8': '20', 'days9': '20', 'loadrate': '5',
        #                'repayMethod': 'ave_capital', 'payTime': '4', 'loanAmount_1': '100000',
        #                'loanPercentage_1': '62.4', 'dep_year1': '10', 'res_rate1': '10', 'dep_year2': '10',
        #                'res_rate2': '10', 'dep_year3': '10', 'depreciationMethod': 'SYD', 'amo_year1': '4',
        #                'amo_year2': '4', 'production1': '100', 'production2': '60', 'production3': '100',
        #                'production4': '3', 'production5': '100', 'production6': '10', 'production7': '20',
        #                'production8': '5', 'production9': '100', 'production10': '140', 'production11': '17',
        #                'production12': '4', 'production13': '2', 'production14': '25', 'production15': '10',
        #                'production16': '10', 'projectcostCounter': '2', 'othercostCounter': '1',
        #                'constInvenstCounter': '0', 'additionalCosts': '{"主要项目":"1000","辅助项目":"1000","土地征用费":"100000"}',
        #                'investmentPlan': '{}', 'LoanPlan': '{"1":{"amount":"100000","percentage":"62.4"}}',
        #                'leftInvestment': '160160', 'leftInvestmentPercentage': '100.0'}
        #
        # data_book = {'projectname':'教材检测','projectplace':'示例','projectind':'示例','input1':'6','input2':'2','productionLoad3':'60','productionLoad4':'80','productionLoad5':'100','productionLoad6':'100','benchmarkyield':'10','benchmark_s_paybackperiod':'5','benchmark_d_paybackperiod':'5.5','install_rate':'10','trans_rate':'0','pre_rate':'10','preup_rate':'10','additionalCostName_1':'主要项目','additionalCostAmount_1':'100','additionalCostPrice_1':'100000','additionalCostEquipment_1':'200','additionalCostName_2':'辅助项目','additionalCostAmount_2':'100','additionalCostPrice_2':'100000','additionalCostEquipment_2':'200','additionalotherCostName_1':'土地征用费','additionalOtherCostType_1':'tudi','additionalOtherCost_1':'2000','additionalotherCostName_2':'开办费','additionalOtherCostType_2':'other','additionalOtherCost_2':'800','days1':'60','days2':'30','days3':'30','days4':'20','days5':'30','days6':'30','days7':'60','days8':'4','days9':'70','loadrate':'5','repayMethod':'ave_capital','payTime':'4','investAmount_1':'3004','investPercentage_1':'50','loanAmount_1':'1202','loanPercentage_1':'40','dep_year1':'10','res_rate1':'10','dep_year2':'10','res_rate2':'10','dep_year3':'10','res_rate3':'0','depreciationMethod':'residualValue','amo_year1':'4','amo_year2':'4','production1':'100','dan1':'','production2':'60','dan2':'','production3':'100','dan3':'','production4':'3','dan4':'','production5':'100','production6':'10','dan6':'','production7':'20','production8':'5','production9':'100','production10':'140','production11':'17','production12':'4','production13':'2','production14':'25','production15':'10','production16':'10','projectcostCounter':'2','othercostCounter':'2','constInvenstCounter':'1','additionalCosts':{"主要项目":"100","辅助项目":"100","土地征用费":"2000","开办费":"800"},'investmentPlan':'{"1":{"amount":"3004","percentage":"50"}}','LoanPlan':'{"1":{"amount":"1202","percentage":"40"},"2":{"amount":"1202","percentage":"40"}}','leftInvestment':'3004','leftInvestmentPercentage':'50.0'}
        # print(data_book)
        # f = Fdata(data_book)
        # latest_file, sheetmap,xspreadsheetmap = f.toCalExcel()

        return jsonify({"message": "Data saved successfully"})
    except Exception as e:
        app.logger.error(f"Error in submit: {str(e)}")
        app.logger.error(traceback.format_exc())
        return jsonify({"error": str(e)}), 500

@app.route('/get_data')
def get_data():
    global latest_file
    global xspreadsheetmap
    global sheetmap
    try:
        if os.path.exists(latest_file):
            s = renderexcel(latest_file,sheetmap,xspreadsheetmap)
            # print(s)
            return s
        else:
            if latest_file == "空.xlsx":
                return jsonify({"message": "请在左侧输入参数"})
            return jsonify({"error": "No data available"}), 404
    except Exception as e:
        app.logger.error(f"Error in get_data: {str(e)}")
        app.logger.error(traceback.format_exc())
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    # 在后台线程中启动 Flask 服务器
    t = threading.Thread(target=start_server)
    t.daemon = True
    t.start()

    # 创建并启动 webview 窗口
    webview.create_window("My Flask App", "http://localhost:5000")
    webview.start()
    # app.run(debug=True,port=5000)
    # webview.create_window("My Flask App", "http://localhost:5000")
    # webview.start()