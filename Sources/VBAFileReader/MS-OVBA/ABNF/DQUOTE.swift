//
//  DQUOTE.swift
//  
//
//  Created by Hugh Bellamy on 28/11/2020.
//

internal struct DQUOTE: ABNFGrammar {
    init(stream: inout ABNFStream) throws {
        try stream.require("\"")
    }
}
