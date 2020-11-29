import XCTest

import WindowsDataTypesTests

var tests = [XCTestCaseEntry]()
tests += OLENativeStream.allTests()
tests += OLEPresentationStream.allTests()
tests += OLEStream.allTests()
tests += TOCEntry.allTests()
XCTMain(tests)
