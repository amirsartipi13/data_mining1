from SqlManager import SqlManager
import ExcelManager
import itertools



def read_fitem_set(excel_name, sheet_name, base):
    info = ExcelManager.read_rows(excel_name, sheet_name, base)
    freq = []
    for row in range(2, len(info)):
        temp = []
        for val in info[row]:
            if val != None:
                temp.append(val)
        freq.append(temp)
    return freq


def make_rule(sql_file, data, minconf):
    sql_manager = SqlManager(sql_file)
    itemSet_count = {}
    rules = []
    lifts = []
    confs = []
    for itemset in data:
        query = 'select count(Descriptions) from transactions2 where '
        temp = ''
        for item in itemset:
            query += 'Descriptions like ' + '"%' + str(item).replace('"', "'") + '%" and '
            temp += str(item) + ','
        query = query[:-4]
        result = sql_manager.crs.execute(query).fetchall()[0][0]
        temp = temp[:-1]
        itemSet_count[temp] = result
    for k in itemSet_count.keys():
        length = len(k.split(','))
        sup_itemset = itemSet_count.get(k)
        if length > 1:
            item_set = k.split(',')
            result = make_sub_set(item_set, length - 1)
            for r in result:
                leftRull = ''
                rightRull = ''
                for left in r[0]:
                    leftRull += left + ','
                for right in r[1]:
                    rightRull += right + ','

                leftRull = leftRull[:-1]
                rightRull = rightRull[:-1]

                sup_left = itemSet_count.get(leftRull, 10000)
                sup_right = itemSet_count.get(rightRull, 10000)

                if float(sup_itemset / sup_left) > float(minconf):
                    conf = float(sup_itemset / sup_left)
                    lift = conf / sup_right
                    rule = leftRull + '--->' + rightRull
                    if rule not in rules:
                        rules.append(rule)
                        lifts.append(lift)
                        confs.append(conf)

                if float(sup_itemset / sup_right) > float(minconf):
                    conf = float(sup_itemset / sup_right)
                    lift = conf / sup_left
                    rule = rightRull + '--->' + leftRull
                    if rule not in rules:
                        rules.append(rule)
                        lifts.append(lift)
                        confs.append(conf)
    return confs, lifts, rules


def make_rule_excel(confs, lifts, rules, excel_name, sheet_name, base_address):
    ExcelManager.create_sheet(excel_name, sheet_name, base_address)
    for i in range(len(rules)):
        temp = str(rules[i]) + "   confidence is :" + str(confs[i]) + "   lifts is :" + str(lifts[i])
        print(temp)
        # ExcelManager.add_row(excel_name, sheet_name,temp,base_address)


def make_sub_set(l, number):
    l1 = list(map(set, itertools.combinations(l, number)))
    result = []
    for i in range(len(l1)):
        result.append((list(l1[i]), [x for x in l if x not in list(l1[i])]))
    return result


if __name__ == '__main__':
    min_cof = 0.6
    fitemset = read_fitem_set('apriori', '0.05', 'outs')
    confs, lifts, rules = make_rule("information.sqlit", fitemset, min_cof)
    make_rule_excel(confs, lifts, rules, 'rule', str(min_cof), '')
    print("finish")
