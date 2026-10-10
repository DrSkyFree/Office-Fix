电脑安装了MS Office，同时安装了WPS，后续更新WPS后出现右键新建的XLSX文件输入内容后保存总是提示兼容性问题，必须另存为才能保存。

<img width="598" height="371" alt="bace7c0717ed0bfd" src="https://github.com/user-attachments/assets/9a772ca1-c1bc-415a-8bc6-cbd6ed6625e1" />

经过排查发现是注册表中新建xlsx键值中模板文件路径与WPS的模板不一致，根本原因是WPS更新每次都会有新的版本号为文件路径，于是就想着换个思路修复，批处理先识别WPS安装路径，然后将模板文件复制到固定目录，接下来再将固定目录下的模板文件路径写入注册表。后续WPS更新，若模板有变化，运行一次批处理会自动复制最新模板到固定目录。使用AI写批处理，经过多次测试是能修复安装office和WPS后，右键新建菜单无新建docx、xlsx、pptx以及由于注册表新建值模板路径不符带来的保存文件兼容性问题。

修复步骤：

1、检测当前系统注册表是否存在Microsoft Office和WPS右键新建，若有则清理,清理后刷新缓存并重启资源管理器;提示用户检查右键新建是否已完全没有.xlsx .pptx .docx .xls .ppt .doc，然后用户输入Y进行行后续，用户输入N，再次检查注册表是否还存在Microsoft Office和WPS右键新建相关值，若有再次清理，没有则直接继续后续工作。

2、检测当前系统是否已安装Microsoft Office和WPS，若安装了分别列出Microsoft Office和WPS版本号，若没有则提示未检测到，并退出。

3、进行修复方案选择

（1）当前系统仅只安装了Microsoft Office则按Microsoft Office 右键新建方案进行修复，并刷新缓存让右键新建生效；

（2）仅只安装WPS时，按wps方案修复，识别WPS路径，复制模板到固定目录，写入注册表进行修复,并重启资源管理器;

 (3)Microsoft Office和WPS均安装时，让用户选择1.Microsoft Office fix new 2.WPS fix new，用户输入1或者2则开始执行修复方案。

4、修复后验证右键新建在注册表是否成功写入，并重启资源管理器。
