//
//  ABNFGrammar.swift
//  
//
//  Created by Hugh Bellamy on 28/11/2020.
//

internal protocol ABNFGrammar {
    init(stream: inout ABNFStream) throws
}
