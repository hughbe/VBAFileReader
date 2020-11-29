import XCTest
import DataStream
import CompoundFileReader
@testable import VBAFileReader

final class VBAFileTests: XCTestCase {
    func testExample() throws {
        do {
            let data = try getData(name: "hughbe/VBA File", fileExtension: "doc")
            let parentFile = try CompoundFile(data: data)
            var rootStorage = parentFile.rootStorage
            let file = try VBAFile(storage: rootStorage.children["Macros"]!)
            XCTAssertEqual(3, file.vbaStorage.modules.count)
            for (_, module) in file.vbaStorage.modules {
                XCTAssertNotNil(module.sourceCode)
            }
        }
    }

    static var allTests = [
        ("testExample", testExample),
    ]
}
