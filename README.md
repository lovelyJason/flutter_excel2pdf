# 介绍

这是一个flutter项目， 本项目本来是做一个excel转换到word/pdf的功能，作为flutter入门项目，会介绍简单的入门知识（不包括安装，环境配置，这是一个比较麻烦的事）

<img width="360" alt="image" src="https://github.com/lovelyJason/flutter_demo/assets/50656459/a4846602-c610-4ed3-ba45-99de462149a8">

运行过程：
![c](https://github.com/lovelyJason/flutter_demo/assets/50656459/ef340df2-11e2-4b9f-8485-ccd8312cf663)


## 开始

### 运行项目

方法很多哈，这里使用vscode+命令行即可

使用vscode，先cmd+shift+p，输入flutter change services, 选中模拟器，然后

```bash
flutter devices # 能看到可用设备,默认包括web的chrome平台
flutter run
```

### 热重载

修改代码后，在终端键入r

### 打包

```bash
# build为安卓apk
flutter build apk

# build macos
flutter config --enable-macos-desktop
flutter create . # 进入项目根目录创建macos对应的文件,此时flutter devices能看到macOS
flutter run -d macos # 运行macos项目
flutter build macos
```

## 包管理器 && 命令行

通过项目下的pubspec.yaml文件来管理依赖

```yaml
dependencies:
  flutter:
    sdk: flutter


  # The following adds the Cupertino Icons font to your application.
  # Use with the CupertinoIcons class for iOS style icons.
  cupertino_icons: ^1.0.2
  # created by jasonhuagn
  file_picker: ^4.2.7
  excel: ^2.0.0
  archive: ^3.3.0
  path_provider: ^2.0.11

```

相关命令

```bash
flutter --version # 检查flutter 版本
flutter pub get # 安装所有依赖, 其实修改pubspec.yaml的时候，vascode中会执行执行这个检查依赖版本是否和fluuter匹配,这个真的很赞
flutter upgrade # 升级sdk
flutter pub add file_picker # 安装依赖，同时更新pubspec.yaml
```

注意：升级了fluttr sdk以后，可能需要重新指定平台flutter config --enable-macos-desktop，然后再flutter create .，flutter pub get,升级完以后，macos会找不到flutter的链接器


那现在我说flutter算入门了，没问题吧？

## flutter的设计哲学

flutter对于涉及到操作系统级别的访问的设计原则是尽可能跨平台且安全，如获取桌面目录，没有提供API，如果有这种需求，对于特定平台，可以通过插件或者凭条通道实现这类功能

## 一些tips

使用ios和macos开发，需要使用CocoaPods，也是一个包管理器，但是是苹果平台的，有时候缺少依赖，需要进入ios或macos目录， 执行pod install

其实在运行flutter run -d macos时，也会自动执行pod install，但不保证不出问题，所以这些还是得懂（苹果平台真的很操蛋QAQ）

### 对于苹果平台，如果pod install报错

> Analyzing dependencies [!] Unable to find a target named `RunnerTests` in project `Runner.xcodeproj`, did find `Runner`.

需要在xcode添加一些配置，或者移除掉podfile（这是一个ruby的语法文件）中的相关代码

### pod install下载问题

> [!] CDN: trunk URL couldn't be downloaded: https://cdn.cocoapods.org/all_pods_versions_c_0_4.txt Response: SSL peer certificate or SSH remote key was not OK

可以尝试将pod的源改为github

编辑`~/.cocoapods/config.yaml`文件，修改为
```yaml
sources:
  - https://github.com/CocoaPods/Specs.git

```
并在Podfile开头添加
`source 'https://github.com/CocoaPods/Specs.git'`

或者cd ~/.cocoapods/repos中克隆源仓库`git clone https://github.com/CocoaPods/Specs.git master`,这个包巨大，要下很长时间

Specs的作用：

> CocoaPods 是一个依赖管理工具，用于在 iOS 和 macOS 开发中集成第三方库。它使用一个名为 Podfile 的文件来定义项目依赖项，然后通过 pod install 命令将这些依赖项下载并集成到项目中。CocoaPods 依赖于一个名为 Specs 的仓库，其中包含了所有已知 CocoaPods 的配置信息。在执行 pod install 命令时，CocoaPods 会从 Specs 仓库中获取最新的规格文件，以确保你安装的依赖项是最新的并且符合你的 Podfile 中的定义。

### file_picker权限问题

使用file_picker的第三方库，用于弹窗选择文件

```dart
var result = await FilePicker.platform.pickFiles(
      type: FileType.custom, // 设置文件类型为自定义类型。
      allowedExtensions: ['xlsx'], // 允许的文件扩展名为xlsx。
    );
```
在macos下报错
> [ERROR:flutter/runtime/dart_vm_initializer.cc(41)] Unhandled Exception: ProcessException: Operation not permitted Command: which osascript

这个问题，我尝试过给dart，flutter，vscode完全的磁盘访问权限，还是报错，最终，我关闭了debug模式下的沙盒

```xml
  <key>com.apple.security.app-sandbox</key>
  <false/>
```

### docx_template库的错误

参照issue: https://github.com/PavelS0/docx_template_dart/issues/53

这个库占位符不是通过{{}}或{}占位符来的，要在wps或者office word设置标签

还有一个大坑就是，这个库依赖的是word文档的开发者工具，为元素设置标签，在wps里叫内容控件。然后，在wps里，mac版的没有开发者工具这一说，也就没办法编辑标签，mac版的office也没有这个功能。还是得自己造轮子吧
