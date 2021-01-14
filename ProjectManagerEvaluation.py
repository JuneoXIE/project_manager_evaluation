# @Author:Xie Ningwei
# @Date:2020-06-28 14:16:30
# @LastModifiedBy:Xie Ningwei
# @Last Modified time:2021-01-14 20:17:58
#-*- coding : utf-8-*-
# coding: utf-8
import easygui as g
import os
import xlrd
import pandas as pd
import numpy as np
import math


class Measure():
    # 初始化方法
    def __init__(self, data_array, weight, benifit_or_cost):
        # 评价指标的值
        self.data_array = data_array
        # 评价人数
        self.num = self.data_array.size
        # 在评价体系中所占权重
        self.weight = weight
        # 评价指标的性质，属于收益型=1还是成本型=0
        self.benifit_or_cost = benifit_or_cost

    def benifit_normalization_model(self):
        self.data_array = self.data_array / np.max(self.data_array)

        # 基准值1:对所有样本降序后位于后10%处的样本值
        index_T_0 = round(self.num * 0.1) - 1
        T_0 = np.sort(self.data_array)[index_T_0]

        # 基准值2：样本指标值的中位数
        T_0_6 = np.median(self.data_array)

        # 基准3：理想值,所有参选样本中该指标的最大值
        T_1 = np.max(self.data_array)

        measure_normalized = []
        for i in range(self.num):
            if (0 <= self.data_array[i] <= T_0):
                measure_normalized.append(0.4 * math.pow(self.data_array[i] / (T_0 + 0.00001), 2))
            elif (T_0 < self.data_array[i] <= T_0_6):
                measure_normalized.append(0.2 * math.pow((self.data_array[i] - T_0) / (T_0_6 - T_0), 0.5) + 0.4)
            elif (T_0_6 < self.data_array[i] <= T_1):
                measure_normalized.append(0.4 * math.pow((self.data_array[i] - T_0_6) / (T_1 - T_0_6), 2) + 0.6)
            else:
                measure_normalized.append(1.0)

        measure_normalized = np.array(measure_normalized)
        return measure_normalized

    def cost_normalization_model(self):
        self.data_array = self.data_array / np.max(self.data_array)

        # 基准值1：；理想值，所有参选样本中该指标的最小值
        C_0 = np.min(self.data_array)

        # 基准值2：样本的指标值的中位数
        C_0_6 = np.median(self.data_array)

        # 基准值3：对所有参选样本进行升序排列，取后10%处样本的指标值
        index_C_0_4 = self.num - round(self.num * 0.1) - 1
        C_0_4 = np.sort(self.data_array)[index_C_0_4]

        measure_normalized = []
        for i in range(self.num):
            if (self.data_array[i] < C_0):
                measure_normalized.append(1.0)
            elif (C_0 <= self.data_array[i] < C_0_6):
                measure_normalized.append(1.0 - 0.4 * math.pow((self.data_array[i] - C_0) / (C_0_6 - C_0), 2))
            elif (C_0_6 <= self.data_array[i] < C_0_4):
                measure_normalized.append(0.2 * math.sqrt((C_0_4 - self.data_array[i]) / (C_0_4 - C_0_6)) + 0.4)
            else:
                measure_normalized.append(1.0 / math.pow(self.data_array[i] - C_0_4 - 0.5 * math.sqrt(10), 2))

        measure_normalized = np.array(measure_normalized)
        return measure_normalized


    def get_raw_measure_data(self):
        return self.data_array

    def get_normalized_result(self):
        # 归一化操作
        if(self.benifit_or_cost):
            self.normalized_result = self.benifit_normalization_model()
        else:
            self.normalized_result = self.cost_normalization_model()
        return self.normalized_result


class ExcelData():
    # 初始化方法
    def __init__(self, data_path, sheet_name = 'Sheet1'):
        # 读入文件路径
        self.data_path = data_path
        # 工作表名称
        self.sheet_name = sheet_name
        # 文件数据
        self.data = xlrd.open_workbook(self.data_path)
        # 工作表
        self.table = self.data.sheet_by_name(self.sheet_name)
        # 索引（姓名 + 指标名称）
        self.keys = self.table.row_values(0)
        # 指标个数
        self.measure_num = 10
        # 经理人数
        self.manager_num = self.table.nrows - 1

        # 指标权重(预先设计)
        self.measure_weights = np.array([0.2, 0.2 / 3.0, 0.2 / 3.0, 0.2 / 3.0, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1])
        # 指标属性(预先设计)
        self.measure_property = np.array([1, 0, 1, 0, 1, 1, 0, 0, 1, 1])

        # 整理指标
        self.measures = []
        for i in range(1, self.measure_num + 1):
            measure_data_array = np.array(self.table.col_values(i)[1:])
            self.measures.append(Measure(measure_data_array, self.measure_weights[i-1], self.measure_property[i-1]))

        # 计算最终分数
        self.final_scores = np.zeros(self.manager_num, dtype=float)
        for index, measure in enumerate(self.measures):
            self.final_scores += self.measure_weights[index] * measure.get_normalized_result()
        self.final_scores *= 100.0


        # 组织输出，并按分数高低排序
        names = self.table.col_values(0)[1:]
        final_scores = self.final_scores.tolist()
        self.output_dict = dict(map(lambda x, y: [x, y], names, final_scores))

        # 根据分数排序
        self.output_dict = sorted(self.output_dict.items(), key = lambda x:x[1], reverse=True)

    def print_info(self):
        print(self.keys)
        print(self.output_dict)

    def get_output_dict(self):
        return self.output_dict



if __name__ == '__main__':
    output_dict = {}
    while(1):
        # 基础界面
        if(not output_dict):
            window = g.buttonbox(msg='请导入Excel格式的产品经理绩效数据数据', title='xxx项目经理评价系统', choices=('导入数据','导出结果','退出'))
        else:
            window = g.buttonbox(msg='绩效数据已导入，请导出结果', title='xxx产品经理评价体系', choices=('导入数据', '导出结果','退出'))

        if(window == '退出'):
            break
        if(window == '导出结果'):
            if(not output_dict):
                g.exceptionbox('错误：请先导入数据！')
                continue
            else:
                try:
                    export_path = g.filesavebox('导出数据',default='*.xlsx')
                    # 将字典列表转换为DataFrame
                    pf = pd.DataFrame(list(output_dict))
                    # 指定标题
                    export_file = pd.ExcelWriter(export_path)
                    pf.to_excel(export_file, encoding='utf-8', index=False, header=['姓名','评价分数'])
                    export_file.save()
                    g.msgbox('分数文件保存成功！',ok_button='退出')
                    break
                except:
                    g.exceptionbox('错误：文件保存失败！')
                    continue

        if(window == '导入数据'):
            data_path = g.fileopenbox('导入数据',default='*.xlsx')
            # 检查文件是否可以打开
            try:
                data = xlrd.open_workbook(data_path)
            except FileExistsError:
                g.exceptionbox('错误：文件不存在！')
                continue
            except:
                g.exceptionbox('错误：文件导入失败！')
                continue

            # 检查文件格式是否符合标准
            try:
                table = data.sheet_by_name('Sheet1')
            except:
                g.exceptionbox('错误：默认应该将数据所在页面命名为Sheet1！')
                continue

            if table.ncols != 11:
                g.exceptionbox('错误：评价指标数量错误，本评价系统要求输入10个评价指标！')
                continue

            # 计算结果
            data = ExcelData(data_path)
            output_dict = data.get_output_dict()
            g.msgbox('数据导入成功！',ok_button='好的')

