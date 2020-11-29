//
//  GUID.swift
//  
//
//  Created by Hugh Bellamy on 28/11/2020.
//

import WindowsDataTypes

/// [MS-OVBA] 2.1.1.5 GUID
/// Specifies a GUID.
/// ABNF syntax:
/// GUID = "{" 8HEXDIG "-" 4HEXDIG "-" 4HEXDIG "-" 4HEXDIG "-" 12HEXDIG "}"
public struct GUID: ABNFGrammar {
    public let value: WindowsDataTypes.GUID
    
    init(stream: inout ABNFStream) throws {
        try stream.require("{")
        let data1: UInt32
        do {
            let stringValue = try stream.readString(count: 8)
            guard let value = UInt32(stringValue, radix: 16) else {
                throw VBAFileError.corrupted
            }
            
            data1 = value
        }
        try stream.require("-")
        let data2: UInt16
        do {
            let stringValue = try stream.readString(count: 4)
            guard let value = UInt16(stringValue, radix: 16) else {
                throw VBAFileError.corrupted
            }
            
            data2 = value
        }
        try stream.require("-")
        let data3: UInt16
        do {
            let stringValue = try stream.readString(count: 4)
            guard let value = UInt16(stringValue, radix: 16) else {
                throw VBAFileError.corrupted
            }
            
            data3 = value
        }
        try stream.require("-")
        
        var data4: [UInt8] = []
        for i in 0..<8 {
            if i == 2 {
                try stream.require("-")
            }

            let stringValue = try stream.readString(count: 2)
            guard let value = UInt8(stringValue, radix: 16) else {
                throw VBAFileError.corrupted
            }
            
            data4.append(value)
        }

        try stream.require("}")
        
        self.value = WindowsDataTypes.GUID(data1, data2, data3, data4)
    }
}
