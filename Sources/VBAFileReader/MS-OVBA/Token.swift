//
//  Token.swift
//  
//
//  Created by Hugh Bellamy on 25/11/2020.
//

import DataStream

public protocol Token {
    func decompress(to: inout [UInt8])
    
    var count: Int { get }
}
