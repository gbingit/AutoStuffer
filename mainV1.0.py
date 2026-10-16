import openpyxl as xl
from datetime import datetime

import subprocess
# subprocess是啥？
# Python调用来跑外部程序的库，这里是要跑osascript，即Mac的程序；在子进程跑，主进程一定是Python3
from typing import List
# 这是啥意思？专门用来注解数据类型的，新版本不需要import，可直接写list[Config]

def to_normalize(text):
    target = str(text)
    return target.lower().replace(' ', '').replace(',', '').replace('-', '').replace(';', '')

def take_out_one_from_dict(dic):
    if dic.keys():
        for key in dic.keys():
            if dic[key] == 0:
                continue
            dic[key] -= 1
            return key
        return None
    else:
        return None

dedicated_group = ['hve', 'reliabilityinregion', 'systempdcupertino', 'liveon', 'wdemc']
higher_priority_groups = ['MSQ - Integration (Factory SW)', 'MSQ - Stress & Stability',
                'Performance-Analysis', 'Performance-Tools'] #(后期要做成可变更
normed_higher_priority_groups = []
for k in higher_priority_groups:
    normed_k = to_normalize(k)
    normed_higher_priority_groups.append(normed_k)

request_ID_mapping = {'requestid': 'systemid'}

def to_dic_head_index(sheet, head_row = 1): #输出表头的index字典（从1开始）
    index_lib = {}
    for index, cell in enumerate(sheet[head_row]):
        if cell.value is not None and cell.value != '':
            index_lib[to_normalize(cell.value)] = index + 1
    return index_lib

def to_compare_stuff_code(astuffCode, bstuffCode):
    result = True
    for idx, l in enumerate(str(astuffCode)):
        if l != 'x':
            if bstuffCode[idx] != l:
                result = False
                break
    return result

class Config(object):
    def __init__(self, name: str, build_date: str, available_qty: int, bstuff_code: str):
        self.name = name
        if build_date:
            try:
                self.build_date = datetime.strptime(build_date, "%m/%d/%Y")
                # .strptime可以解析什么类型的数据？
                # 后面的第二个参数是说明按什么样的格式解读数据
            except ValueError:
                self.build_date = None
        else:
            self.build_date = None
        self.available_qty = available_qty
        self.bstuff_code = bstuff_code

    @property
    def is_available(self) -> bool:
        return self.available_qty > 0

    def take_out(self, num:int):
        amount_to_take = min(num,self.available_qty)
        self.available_qty -= amount_to_take
        return amount_to_take

    def __repr__(self):
    # __repr__这是什么
    # """打印对象实例时输出自定义信息，替代默认的内存地址字符串，方便调试查看对象属性"""
        return f"Config(name='{self.name}', build_daye={self.build_date}, qty={self.available_qty}, bstuff_code='{self.bstuff_code}', is_available={self.is_available})"

def take_out_from_Config_series(config_list:list[Config], qty:int):
    result_dict = {}
    remain_qty = qty
    for config in config_list:
        if remain_qty <= 0:
            break
        actual_qty = config.take_out(remain_qty)
        #这里会返回.take_out这个function里返回的amount_to_take吗？
        #yes，但没东西接着就不会显示出来
        remain_qty -= actual_qty
        result_dict[config.name] = actual_qty
    return remain_qty, result_dict

class ExclusiveConfig(Config):
    def __init__(self, name: str, build_date: datetime | None, bstuff_code: str):
            # 给父类传空字符串，骗过父类的strptime解析
            super().__init__(name, build_date="", available_qty=0, bstuff_code=bstuff_code)
            # 回到子类，覆盖掉父类生成的临时 build_date
            self.build_date = build_date
            # 因为这里是子到父再到子的顺序，父已经定义为接受str，再在初始化的时候转化为datetime；
            # convert_configs_to_exclusive 规定好了传到子里的参数是在父里已经定义好的datetime类型，所以子必须接受datetime
            # 子接受datetime，但父只能接受str，所以要用一个空str骗父先全初始化了，然后后面再单独初始化一下.datetime

    def __repr__(self):
        return f"ExclusiveConfig(name='{self.name}', qty={self.available_qty})"


def mac_get_exclusive_name_list() -> List[str]:
    """
    macOS系统弹窗，用户输入name，多个名称用英文逗号分隔；取消返回空列表
    注意：原生display dialog输入框为单行，不支持真正换行
    """
    # 重点：text returned of (...) 直接提取输入文本，不会输出 button returned 前缀
    applescript = '''
text returned of (display dialog "输入需要转为ExclusiveConfig的name，多个名称用英文逗号分隔" with title "标记为专属配置" default answer "")
'''
    proc = subprocess.run(
        ["osascript", "-e", applescript],
        stdout=subprocess.PIPE, #？osascript 正常返回的用户输入 → 通过管道给到proc.stdout
        stderr=subprocess.PIPE, #？osascript 的系统警告、报错 → 通过管道给到proc.stderr
        text=True #？全部转为普通字符串，Python 代码读取proc.stdout拿到用户输入
    )

    # returncode !=0：用户按取消/ESC，osascript报错 -128
    if proc.returncode != 0:
        print("DEBUG stderr错误详情：", proc.stderr)
        return []

    output = proc.stdout.strip()
    names = [line.strip() for line in output.split(",") if line.strip()]
    return names

def convert_configs_to_exclusive(original_configs: List[Config], target_names: List[str]) -> List[ExclusiveConfig]:
    """
    传入原始Config列表 + 用户输入的name列表
    将匹配name的Config转换生成ExclusiveConfig，复用原有 build_date / bstuff_code，qty强制置0
    不在original_configs中的name直接跳过并打印警告
    """
    exclusive_list: List[ExclusiveConfig] = []
    # 构建查找字典，方便快速匹配 name → Config
    config_map = {cfg.name: cfg for cfg in original_configs}
    print(config_map)
    # 这里的cfg.name是什么？只是cfg的name属性，而后面的cfg是其整体

    for name in target_names:
        if name not in config_map:
            print(f"[WARN] 找不到name={name} 的原始Config，跳过")
            continue
        src_cfg = config_map[name]
        # 复用原Config的字段，生成ExclusiveConfig
        ex_cfg = ExclusiveConfig(
            name=src_cfg.name,
            build_date=src_cfg.build_date,
            bstuff_code=src_cfg.bstuff_code
        )
        exclusive_list.append(ex_cfg)
    return exclusive_list

class Allocation(object):
    def __init__(self, group, astuff_code, request_id, approved_qty, approved_unit_details, stuff_qty, approved_build_phase, approved_ship_out_priority, ship_to_location):
        self.group = group
        self.astuff_code = astuff_code
        self.request_id = request_id
        self.approved_qty = approved_qty
        #如果approved_qty = 0?
        self.is_effective = self.approved_qty > 0
        self.approved_unit_details = approved_unit_details
        self.requirement_qty = self._calc_requirement_qty(astuff_code)
        #这里的requirement_qty是个计算属性？no，这里的写法是普通实例属性，只在init的时候计算一次，后面就算改了astuff_code也不跟着变化
        self.is_handpick = any(k in to_normalize(str(self.approved_unit_details)) for k in ("handpick", "cherrypick", "select")) if self.approved_unit_details else False
        #any之中的元组("handpick", "cherrypick", "selected")只要拿到一个True就直接返回True，后面不再执行
        self.has_exclusive_cfg = any(k in to_normalize(self.group) for k in dedicated_group)
        self.stuff_qty = stuff_qty
        self.approved_build_phase = approved_build_phase
        self.approved_ship_out_priority = approved_ship_out_priority
        self.ship_to_location = ship_to_location
        self.is_domestic = any(k in to_normalize(self.ship_to_location) for k in ('beijing', 'stayatcm', 'china')) if self.ship_to_location else False
        self.is_higher_priority_group = any(k in to_normalize(self.group) for k in normed_higher_priority_groups) if self.group else False
        cond_0 = (
                'hve' in to_normalize(self.group)
                or 'reliabilityinregion' in to_normalize(self.group)
                or (len(self.astuff_code) >= 2 and self.astuff_code[1] == 'N'))
        if cond_0:
            self.priority = 0
        elif to_normalize(self.approved_ship_out_priority) == 'urgent':
            if self.is_higher_priority_group:
                if self.is_domestic:
                    self.priority = 2
                else:
                    self.priority = 1
            else:
                if self.is_domestic:
                    self.priority = 3
                else:
                    self.priority = 2
        else:
            if self.is_higher_priority_group:
                if self.is_domestic:
                    self.priority = 4
                else:
                    self.priority = 3
            else:
                if self.is_domestic:
                    self.priority = 5
                else:
                    self.priority = 4

    @property
    def is_stuffed(self):
        if self.is_handpick or self.has_exclusive_cfg:
            return True
        else:
            if self.stuff_qty < self.approved_qty:
                return False
            else:
                return True

    @staticmethod
    # 静态方法，class可直接调用，且不用参数self（即不用打_calc_requirement_qty()）
    # Why要特意把这个func写出静态？这个逻辑属于 Allocation 的附属工具，但是计算不需要实例对象的任何数据，只靠传入参数，所以做成静态方法。
    def _calc_requirement_qty(code):
        """统计字符串中非 'x' 字符的个数"""
        count = 0
        for char in code:
            if char != 'x':
                count += 1
        return count

    def __repr__(self):
        return f"{self.__class__.__name__}({self.__dict__!r})"

class SubAllocation(object):
    _request_ID_counter = {}
    # _前置表示这个属于类私有变量，不属于某个实例，也不推荐外部直接访问
    def __init__(self, request_ID, stuffed_cfg, niche):
        self.request_ID = request_ID
        if request_ID not in self._request_ID_counter:
            self._request_ID_counter[request_ID] = 0
        self._request_ID_counter[request_ID] += 1
        self.num = self._request_ID_counter[request_ID]
        self.stuffed_cfg = stuffed_cfg
        self.niche = niche

    @property
    def is_stuffed(self):
        return self.stuffed_cfg is not None

    @classmethod
    def reset_counter(cls):
        """处理完一张表格，重置计数器，下一轮从1重新计数"""
        cls._request_ID_counter.clear()

    def __repr__(self):
        return f"{self.__class__.__name__}({self.__dict__!r})"


if __name__ == "__main__":
    wb = xl.load_workbook('/Users/edison/Desktop/AutoStufferMac/PWZM EVT(c) S2F.xlsx', data_only=True)
    allocation_sheet = wb['Allocations']
    s2f_sheet = wb['S2F']
    bm_sheet = wb['Build Matrix']
    allocation_head_dic = to_dic_head_index(allocation_sheet)
    s2f_head_dic = to_dic_head_index(s2f_sheet)
    bm_head_dic = to_dic_head_index(bm_sheet)
    allocation_list = []
    config_list = []
    sub_allocation_list = []
    exclusive_allocation_list = []
    new_append_col = allocation_sheet.max_column + 1


    for row in range(2, allocation_sheet.max_row + 1):
        col_group = allocation_head_dic["group"]
        col_astuff_code = allocation_head_dic["astuffconfig"]
        col_req_id = allocation_head_dic["requestid"]
        col_approved_qty = allocation_head_dic["approvedqty"]
        col_approved_unit_details = allocation_head_dic["approvedunitdetails"]
        col_approved_build_phase = allocation_head_dic["approvedbuildphase"]
        col_approved_ship_out_priority = allocation_head_dic["approvedshipoutpriority"]
        col_ship_to_location = allocation_head_dic["shiptolocation"]

        val_group = allocation_sheet.cell(row=row, column=col_group).value
        val_astuff_code = allocation_sheet.cell(row=row, column=col_astuff_code).value
        val_req_id = allocation_sheet.cell(row=row, column=col_req_id).value
        val_approved_qty = allocation_sheet.cell(row=row, column=col_approved_qty).value
        val_approved_unit_details = allocation_sheet.cell(row=row, column=col_approved_unit_details).value
        val_approved_build_phase = allocation_sheet.cell(row=row, column=col_approved_build_phase).value
        val_approved_ship_out_priority = allocation_sheet.cell(row=row, column=col_approved_ship_out_priority).value
        val_ship_to_location = allocation_sheet.cell(row=row, column=col_ship_to_location).value

        if val_group is not None and val_astuff_code is not None and val_req_id is not None and val_approved_qty is not None:
            alloc_obj = Allocation(
                group=val_group,
                astuff_code=val_astuff_code,
                request_id=val_req_id,
                approved_qty=val_approved_qty,
                approved_unit_details=val_approved_unit_details,
                stuff_qty=0,
                approved_build_phase=val_approved_build_phase,
                approved_ship_out_priority=val_approved_ship_out_priority,
                ship_to_location=val_ship_to_location
            )
            allocation_list.append(alloc_obj)
            allocation_sheet.cell(row, new_append_col).value = alloc_obj.priority
            if alloc_obj.is_stuffed:
                exclusive_allocation_list.append(alloc_obj.request_id)
    print(f'exclusive_allocation_list{exclusive_allocation_list}')


    # 测试：打印第一个对象验证
    if allocation_list:
        first = allocation_list[1]
        print(first)

    for row in range(2, s2f_sheet.max_row + 1):
        col_request_id = s2f_head_dic.get(request_ID_mapping['requestid'])
        col_config = s2f_head_dic.get('config')
        val_request_id = s2f_sheet.cell(row=row, column=col_request_id).value
        target_cell = s2f_sheet.cell(row=row, column=col_config)

        if val_request_id:
            sub_allocation = SubAllocation(request_ID=val_request_id, stuffed_cfg=None, niche=target_cell)
            sub_allocation_list.append(sub_allocation)

    print(sub_allocation_list[10:15])

    for row in range(2, bm_sheet.max_row + 1):
        col_name = bm_head_dic["name"]
        col_build_date = bm_head_dic["builddate"]
        col_available_qty = bm_head_dic["targetoutput"]
        col_bstuff_code = bm_head_dic["bstuffconfig"]
        col_build_phase = bm_head_dic["buildphase"]

        val_name = bm_sheet.cell(row=row, column=col_name).value
        val_build_date = bm_sheet.cell(row=row, column=col_build_date).value
        val_available_qty = int(bm_sheet.cell(row=row, column=col_available_qty).value)
        val_bstuff_code = bm_sheet.cell(row=row, column=col_bstuff_code).value
        val_build_phase = bm_sheet.cell(row=row, column=col_build_phase).value

        if (val_name is not None and val_build_date is not None and val_build_phase != "U" and
                val_available_qty is not None and val_bstuff_code is not None):
            config_obj = Config(
                name=val_name,
                build_date=val_build_date,
                available_qty=val_available_qty,
                bstuff_code=val_bstuff_code
            )
            config_list.append(config_obj)

    if config_list:
        example = config_list[121]
        print(example)

    #怎么确定config_list里哪些config是Exclusive ones？
    exclusive_list = ['EVTc-Mini1PDID1-A-N-36M', 'EVTc-Mini4PDID4-A-B-64M', 'EVTc-Mini5PDID5-A-B-128M'
                      , 'EVTc-Mini3PDID6-A-B-48M', 'Shmoo', 'Live']
    selected_names = mac_get_exclusive_name_list()
    if not selected_names:
        print("用户取消输入，无ExclusiveConfig生成")
    else:
        # 2.匹配原始config，转换为ExclusiveConfig列表
        exclusive_result = convert_configs_to_exclusive(config_list, selected_names)
        print("\n转换完成，得到ExclusiveConfig列表：")
        for item in exclusive_result:
            print(item, f"is_available={item.is_available}")

    sorted_allocation_list = sorted(allocation_list, key=lambda x: (x.priority, -x.requirement_qty))
    # print(sorted_allocation_list[:5])
    unfulfilled_allocation_list = []
    for allocation in sorted_allocation_list:
        if allocation.is_stuffed:
            continue
        else:
            match_count = 0
            matched_cfg_list = []
            for config in config_list:
                if config.is_available and to_compare_stuff_code(allocation.astuff_code, config.bstuff_code):
                    matched_cfg_list.append(config)
            if matched_cfg_list:
                sorted_cfg_list = sorted(matched_cfg_list, key=lambda x: x.build_date)
                remain, stuffed_configs_dict = take_out_from_Config_series(sorted_cfg_list, allocation.approved_qty)
                # !!这只是从config里拿出来，但还有allocation被stuff的function
                allocation.stuff_qty += (allocation.approved_qty - remain)
                if remain > 0:
                    unfulfilled_allocation_list.append(allocation.request_id)
                    print(f'there are still {remain}x {allocation.request_id} that have no configs to stuff, {stuffed_configs_dict}')
                else:
                    print(f"{allocation.request_id}'s is_stuffed is {allocation.is_stuffed}, stuffed by {stuffed_configs_dict}")

            # 连接Allocation和SubAllocation using request_id
                for sub_allocation in sub_allocation_list:
                    if match_count >= allocation.approved_qty:
                        break
                    if not sub_allocation.is_stuffed and sub_allocation.request_ID == allocation.request_id:
                        match_count += 1
                        cfg_to_be_stuffed = take_out_one_from_dict(stuffed_configs_dict)
                        if cfg_to_be_stuffed is None:
                            sub_allocation.niche.value = 'empty stuffed_configs_dict'
                            break
                        else:
                            sub_allocation.stuffed_cfg = cfg_to_be_stuffed
                    sub_allocation.niche.value = sub_allocation.stuffed_cfg if sub_allocation.stuffed_cfg is not None else f'{sub_allocation.request_ID}'


            # 会存在没有match的config的allocation吗？

    print(len(unfulfilled_allocation_list), unfulfilled_allocation_list)
    wb.save('/Users/edison/Desktop/AutoStufferMac/PWZM EVT(c) S2F.test.xlsx')









