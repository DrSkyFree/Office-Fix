#修复安装office和WPS后，右键新建菜单无新建docx、xlsx、pptx#
电脑安装了MS Office，同时安装了WPS，后续更新WPS后出现右键新建的XLSX文件输入内容后保存总是提示兼容性问题，必须另存为才能保存。
<img width="598" height="371" alt="bace7c0717ed0bfd" src="https://github.com/user-attachments/assets/9a772ca1-c1bc-415a-8bc6-cbd6ed6625e1" />
经过排查发现是注册表中新建xlsx键值中模板文件路径与WPS的模板不一致，根本原因是WPS更新每次都会有新的版本号为文件路径，于是就想着换个思路修复，批处理先识别WPS安装路径，然后将模板文件复制到固定目录，接下来再将固定目录下的模板文件路径写入注册表。后续WPS更新，若模板有变化，运行一次批处理会自动复制最新模板到固定目录。方案有了，就交给AI来写批处理了，经过多次测试是能修复安装office和WPS后，右键新建菜单无新建docx、xlsx、pptx。
修复步骤：
1、检测当前系统右键新建并清理；
2、检测当前系统安装MS Office和WPS
3、当前系统仅安装了MS Office则按MS Office关联进行修复，若仅安装WPS或二者皆安装，就选择WPS修复方案，识别WPS路径，复制模板到固定目录，写入注册表进行修复；
4、验证右键新建在注册表是否成功写入，并重启资源管理器。
