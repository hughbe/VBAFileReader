//
//  File.swift
//  
//
//  Created by Hugh Bellamy on 28/11/2020.
//

import DataStream

internal struct ABNFStream {
    private var dataStream: DataStream
    
    public var position: Int {
        get { dataStream.position }
        set { dataStream.position = newValue }
    }
    
    public var count: Int { dataStream.count }
    
    public init(dataStream: inout DataStream) {
        self.dataStream = dataStream
    }
    
    public mutating func nextIndex(string: String) -> Int? {
        let startPosition = dataStream.position
        defer {
            dataStream.position = startPosition
        }

        while !peek(string: string) {
            guard dataStream.position + 1 < dataStream.count else {
                return -1
            }

            dataStream.position += 1
        }
        
        return dataStream.position
    }
    
    public mutating func readString(count: Int) throws -> String {
        return try dataStream.readString(count: count, encoding: .ascii)!
    }
    
    public mutating func peekString(count: Int) throws -> String {
        return try dataStream.peekString(count: count, encoding: .ascii)!
    }
    
    public mutating func readByte() throws -> UInt8 {
        return try dataStream.read()
    }
    
    public mutating func peekByte() throws -> UInt8 {
        return try dataStream.peek()
    }
    
    public mutating func read<T>(_ grammar: T.Type) throws -> T where T: ABNFGrammar {
        return try T(stream: &self)
    }
    
    public mutating func read<T>() throws -> T where T: ABNFGrammar {
        return try T(stream: &self)
    }
    
    public mutating func require<T>(_ grammar: T.Type) throws where T: ABNFGrammar {
        let _ = try T(stream: &self)
    }
    
    public mutating func require(_ string: String) throws {
        if !peek(string: string) {
            throw VBAFileError.corrupted
        }
        
        position += string.count
    }

    public mutating func peek<T>(grammar: T.Type) -> T? where T: ABNFGrammar {
        let position = dataStream.position
        defer {
            dataStream.position = position
        }
        
        return try? T(stream: &self)
    }
    
    public mutating func peek(string: String) -> Bool {
        guard position + string.count <= count else {
            return false
        }

        let startPosition = position
        defer {
            position = startPosition
        }
        
        for i in string.indices {
            guard try! dataStream.read() as UInt8 == string[i].asciiValue else {
                return false
            }
        }

        return true
    }
}
