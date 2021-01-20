# 使用步骤:
# 先要确保运行运行脚本的机器可以通过powershell访问AD，并可以通过copy命令将文件从AD拷贝到本地
# 先运行本python脚本，生成powershell脚本文件
# 打开powershell脚本文件，复制其内容
# 打开powershell命令窗口，将脚本内容粘贴到命令窗口
# 开启AD powershell远程访问可参考：https://www.cnblogs.com/sparkdev/p/7200004.html
# powershell脚本命令执行完成后，本机存放python脚本目录下获得AD用户csv文件
# 再一次运行本python脚本，此时，python脚本会做以下动作：
# 1. 根据输入的关键字，在AD用户csv文件中提取包含关键字用户组的用户信息，并将提取信息后的AD用户csv文件转换成可以录入ISE的csv格式.
# 2. 检查ISE上是否已经创建所需的用户组，如果没有，则会询问是否需要自动创建
# 3. 检查ISE上是否有AD中不存在的用户，如果有，则会询问是否需要自动删除
# 最后，手动导入脚本生成的ISE csv文件

import requests
from requests.auth import HTTPBasicAuth
import sys
import os
import pandas as pd
import json
import re
import pprint


# 0. 生成ps脚本，并写入登陆命令
# 不知道什么原因，生成的ps1脚本不能直接点击执行。需要复制命令行，贴到powershell窗口执行

# 本脚本所在目录
path = sys.path[0] + '\\'
# powershell文件名
ps_name = 'ad_users.ps1'

# 从AD获取的用户信息csv
ad_csv_name = 'ad_users.csv'
# AD转换后需要上传到ISE的用户信息csv
ise_csv_name = 'ad_to_ise_users.csv'

# ISE登陆信息
# ip或fqdn
ise_ipadd = '198.18.133.27'
ise_username = 'ersadmin'
ise_password = 'Cisco123'
# ISE API call要求的必要的头部信息，这部分不要改动
ise_headers = {
				'Content-Type': 'application/json', 
				'Accept': 'application/json'
				}

# AD登陆信息
ad_ipadd = '198.18.133.1'
ad_username = 'administrator'
ad_password = 'C1sco12345'

ad_name_on_ise = 'All_AD_Join_Points'

# 定义用户组所包含的关键字，大小写不敏感
keyword = 'grade'


#############################################################################################

def generate_ps_script(username = 'administrator', password = 'C1sco12345', ad_ip = '198.18.133.1'):
	with open(path + 'ad_users.ps1','w+',encoding="utf-8") as f:
	
		login = f'''
$Username = '{username}'
$Password = '{password}'
$pass = ConvertTo-SecureString -AsPlainText $Password -Force
$Cred = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$pass
Enter-PSSession -ComputerName {ad_ip} -Credential ($Cred)
'''
		
		part_1 = '''get-aduser -filter * -SearchBase 'DC=dcloud,DC=cisco,DC=com' -properties SamAccountName, memberof, GivenName, Surname | Where-Object {$_.Enabled -eq $True} | select SamAccountName, @{name="First Name";expression={$_.GivenName -join " "}}, @{name="Last Name";expression={$_.Surname -join " "}}, @{name="memberof";expression={$_.memberof -join ";"}} '''
		
		part_2 = f'''| export-csv "c:\{ad_csv_name}" -notypeinformation -Encoding UTF8'''
		
		ad_users_csv = part_1 + part_2
		
		
		commands = (login, f'Import-Module ActiveDirectory\n', ad_users_csv, f'\nexit', f'\ncopy \\\{ad_ip}\c$\{ad_csv_name} {path} \n')
		for command in commands:
			f.write(command)
		
		print('Powershell script is ready! \n')

generate_ps_script(ad_username, ad_password, ad_ipadd)

# 2. 解析从ad获取的用户csv
# 用正则表达式从memberof中提取感兴趣组
def get_group_from_mf(mf):
	
	# 关键字如果在memberof最后位置，正则无法匹配末尾空字符，因此在末尾添加上空格
	mf = mf + ' '
	
	pattern = re.compile(f'cn=([^=]*{keyword}.*?)[,\s]', re.I)
	
	group_name = pattern.findall(mf)
	
	if group_name != []:
	
		return group_name[0].upper()

def get_interested_group_csv():
	
	df = pd.read_csv(path + ad_csv_name)
	
	# 将所有memberof列带有空值的行都去掉, 空值无法进行函数运算，所以必须先清洗掉
	new_df = df.dropna(axis=0, subset=['memberof'])
	
	# 使用map函数将memberof列中的group名提取出来，并替换到现有的memberof列中
	new_df['memberof'] = new_df['memberof'].map(get_group_from_mf)
	
	# 将memberof列更名为Group
	new_df.rename(columns = {'memberof':'User Identity Groups', 'SamAccountName':'User Name'}, inplace=True)
	
	# 将所有Group列带有空值的行都去掉
	new_table = new_df.dropna(axis=0, subset=['User Identity Groups'])
	
	# 补齐ISE所需columns
	new_table = new_table.join(pd.DataFrame(
	{
		'Is Password Encrypted(True/False)':'False',
		'Enable User(Yes/No)':'Yes',
		'Change Password on Next Login(Yes/No)': 'No',
		'Password ID Store': ad_name_on_ise,
		'Email': '',
		'User Details': '',
		'Password': '',
		'Enable Password': '',
		'Expiry Date(MM/dd/yyyy)': '',
		'Is Enable Password Encrypted(True/False)': ''
	}, index = new_table.index
	))
	
	# 将新矩阵保存到csv文件
	new_table.to_csv(path + ise_csv_name, index=False)
	
	return new_table

# 用户创建ISE用户组的函数
def ise_create_user_group(gp_name):

	url = f'https://{ise_ipadd}:9060/ers/config/identitygroup'


	body = {
	  "IdentityGroup" : {
		"id" : "id",
		"name" : f"{gp_name}",
		"description" : f"{gp_name}_description",
		"parent" : "parent"
	  }
	}

	body = json.dumps(body)
		
	r_c = requests.post(url, auth = HTTPBasicAuth(ise_username,ise_password), headers = ise_headers, data = body, verify = False)

	response = r_c
	
	if response.status_code == 201:
		print(f'{gp_name} Created!')
		
	else:
		pprint.pprint(response.json())


# 用于获得ISE用户组的函数
def get_ise_usg(group_name = ''):

	if group_name == '':
		url = f'https://{ise_ipadd}:9060/ers/config/identitygroup?size=100&page=1'
	else:
	## 可以用filter语句过请求包含关键字的用户组
		url = f'https://{ise_ipadd}:9060/ers/config/identitygroup/?filter=name.CONTAINS.{group_name}&page=1'
	
	resp_groups = []
	
	# 第1个url页面内容
	temp_SearchResult = get_response(url)['SearchResult']
	print(url)
	
	# 第1(或n)个url页面内容有nextPage则进入循环
	while temp_SearchResult.get('nextPage'):
		# 将第1页的用户信息累加到groups
		resp_groups += temp_SearchResult['resources']
		
		# 将下一个url更新到循环外
		url = temp_SearchResult['nextPage']['href']
		# 将下一个url的内容更新到循环外
		temp_SearchResult = get_response(url)['SearchResult']
		print(url)
		
	# 最后一页没有nextPage,因此跳出循环，此时再将最后一页的用户累加到resp_groups
	resp_groups += get_response(url)['SearchResult']['resources']
	
	groups = []
	
	for group in resp_groups:
		groups.append(group['name'])
	
	return groups


# 将页面获得的用户列表生成为用户及其id对应的字典
def gen_user_dict(ul):
	user_dict = {}
	
	for i in ul:		
		user_dict.setdefault(i['name'],i['id'])
	
	return user_dict

# 获取页面内容
def get_response(url):
	
	r_c = requests.get(url, auth = HTTPBasicAuth(ise_username, ise_password), headers = ise_headers, verify = False)
	
	response = r_c.json()
	
	return response

# 获取ISE用户的函数
def get_ise_users(group_name = ''):
	
	if group_name == '':
		url = f'https://{ise_ipadd}:9060/ers/config/internaluser?size=100&page=1'
		
	else:
	## 可以用filter语句过请求在某个用户组的用户
		url = f'https://{ise_ipadd}:9060/ers/config/internaluser/?filter=identityGroup.CONTAINS.{group_name}&page=1'
	
	user_list = []
	
	# 第1个url页面内容
	temp_SearchResult = get_response(url)['SearchResult']
	print(url)
	
	# 第1(或n)个url页面内容有nextPage则进入循环
	while temp_SearchResult.get('nextPage'):
		# 将第1页的用户信息累加到user_list
		user_list += temp_SearchResult['resources']
		
		# 将下一个url更新到循环外
		url = temp_SearchResult['nextPage']['href']
		# 将下一个url的内容更新到循环外
		temp_SearchResult = get_response(url)['SearchResult']
		print(url)
		
	# 最后一页没有nextPage,因此跳出循环，此时再将最后一页的用户累加到user_list
	user_list += get_response(url)['SearchResult']['resources']
	
	return gen_user_dict(user_list)


# 用于删除ISE用户的函数
def delete_ise_users(user_dict):
	
	ul = list(user_dict)
	for i in ul:
		user_id = user_dict[i]
	
		url = f'https://{ise_ipadd}:9060/ers/config/internaluser/{user_id}'
			
		r_c = requests.delete(url, auth = HTTPBasicAuth(ise_username, ise_password), headers = ise_headers, verify = False)
		
		response = r_c
		
		if response.status_code == 204:
			print(i,' has been deleted \n')
		else:
			pprint.pprint(response.json())

# 用于检查AD用户组是否都存在于ISE的函数
def check_groups_in_ise():
	ad_group_set = set(map(lambda x:x.upper(), get_interested_group_csv()['User Identity Groups'].tolist()))
	ise_group_set = set(map(lambda x:x.upper(), get_ise_usg(keyword)))
	ise_missing_group_set = ad_group_set - ise_group_set
	
	if ise_missing_group_set > set():
		print('Group ', str(ise_missing_group_set), ' are missing! Do you want to create them on ISE? \n')
		next_step = input('Y/N: ')
		if next_step.isalpha():
			if next_step.lower() == 'y':
				for i in list(ise_missing_group_set):
					ise_create_user_group(i)
			else:
				print('Plesae make sure you have relevant groups on ISE, before imporing users. \n')


# 检查ISE上是否有AD中不存在的用户的函数
def check_users_in_ise():
	ise_users = get_ise_users()
	ad_user_set = set(get_interested_group_csv()['User Name'].tolist())
	ise_user_set = set(ise_users.keys())
	
	ad_missing_user_set = ise_user_set - ad_user_set
	if ad_missing_user_set > set():
		print('User', str(ad_missing_user_set), ' are missing on AD! Do you want to delete them on ISE? \n')
		next_step = input('Y/N: ')
		if next_step.isalpha():
			if next_step.lower() == 'y':
				for user in ad_missing_user_set:
					delete_ise_users({user:ise_users[user]})
						

if ad_csv_name not in os.listdir(sys.path[0]):
	print(f'{ad_csv_name} not found! \n')
else:
	get_interested_group_csv()
	check_groups_in_ise()
	check_users_in_ise()
	print(f'{ise_csv_name} is generated, please import it to ISE! \n')

