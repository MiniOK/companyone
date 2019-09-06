import re
import warnings

left = """left CCA-IMT(mm)近段0.8中段0.7远段0.8斑块（单位mm，空缺为正常）数量（1=单发，2=多发）2最长者长度3.0最大者厚度1.1形态（1=规则型，2=不规则型）1是否溃疡型（0=否，1=是）0质地（A1=均质低回声，A2=均质等回声，A3=均质强回声，B=不均质）A3管腔直径狭窄率%狭窄部位
"""




try:
    re_text = re.search("(大者长度|最大者长|最长者长度|最大者长度|度)(.*?)(mm|最大者厚度|最大厚度)", left).group(2)
    print(re_text)
except Exception as e:
    warnings.warn(e)

