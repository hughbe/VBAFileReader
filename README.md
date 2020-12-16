# VBAFileReader

Swift definitions for structures, enumerations and functions defined in [MS-OVBA](https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-ovba/)

## Example Usage

Add the following line to your project's SwiftPM dependencies:
```swift
.package(url: "https://github.com/hughbe/VBAReader", from: "1.0.0"),
```

```swift
import CompoundFileReader
import MsgReader

let data = Data(contentsOfFile: "<path-to-file>.doc")!
let parentFile = try CompoundFile(data: data)
var rootStorage = parentFile.rootStorage
let file = try VBAFile(storage: rootStorage.children["Macros"]!)
for module in file.vbaStorage.modules {
    print(module.sourceCode)
}
```
