//
//  VBFrame.swift
//  
//
//  Created by Hugh Bellamy on 28/11/2020.
//

import DataStream

/// [MS-OVBA] 2.3.5 VBFrame Stream: Designer Information
/// The VBFrame stream specifies the extended property values of a designer.
/// This stream is an array of bytes that specifies the extended property values of a designer module. MUST contain MBCS characters encoded
/// using the code page specified in PROJECTCODEPAGE (section 2.3.4.2.1.4).
/// Property values of the designer are set at design-time. Property values are used at run-time as specified to initialize the designer. For example,
/// a designer can be used at run time to display data to and accept data from a user and the following properties could be used to determine
/// the location of the designer.
/// ABNF syntax:
/// VBFrameText = "VERSION 5.00" NWLN
///  "Begin" 1*WSP DesignerCLSID 1*WSP DesignerName *WSP NWLN
///  DesignerProperties "End" NWLN
/// DesignerCLSID = GUID
/// DesignerName = ModuleIdentifier
/// <DesignerCLSID>: Specifies the class identifier (CLSID) of the designer. The Automation type library that contains the designer specified
/// MUST be referenced with a REFERENCECONTROL (section 2.3.4.2.2.3). The value "{C62A69F0-16DC-11CE-9E98-00AA00574A4F}"
/// specifies the designer is an Office Form ActiveX control specified in [MS-OFORMS].
/// <DesignerName>: Specifies the name of the designer module associated with the properties.
public struct VBFrame {
    public let designerCLSID: GUID
    public let designerName: ModuleIdentifier
    public let designerProperties: DesignerProperties
    
    public init(dataStream: inout DataStream, count: Int) throws {
        let startPosition = dataStream.position

        var stream = ABNFStream(dataStream: &dataStream)
        
        try stream.require("VERSION 5.00")
        try stream.require(NWLN.self)
        try stream.require("Begin")
        try stream.require(WSP.self)
        self.designerCLSID = try stream.read(GUID.self)
        try stream.require(WSP.self)
        self.designerName = try stream.read(ModuleIdentifier.self)
        try stream.require(WSP.self)
        try stream.require(NWLN.self)
        
        self.designerProperties = try stream.read(DesignerProperties.self)
        try stream.require("End")
        try stream.require(NWLN.self)
        
        guard stream.position - startPosition == count else {
            throw VBAFileError.corrupted
        }
    }
    
    /// [MS-OVBA] 2.3.5.1 DesignerProperties
    /// Specifies the VBA-specific extended properties of a designer.
    /// ABNF syntax:
    /// DesignerProperties = [ *WSP DesignerCaption *WSP [ Comment ] NWLN ]
    ///  [ *WSP DesignerHeight *WSP [ Comment ] NWLN ]
    ///  [ *WSP DesignerLeft *WSP [ Comment ] NWLN ]
    ///  [ *WSP DesignerTop *WSP [ Comment ] NWLN ]
    ///  [ *WSP DesignerWidth *WSP [ Comment ] NWLN ]
    ///  [ *WSP DesignerEnabled *WSP [ Comment ] NWLN ]
    ///  [ *WSP DesignerHelpContextId *WSP [ Comment ] NWLN ]
    ///  [ *WSP DesignerRTL *WSP [ Comment ] NWLN ]
    ///  [ *WSP DesignerShowModal *WSP [ Comment ] NWLN ]
    ///  [ *WSP DesignerStartupPosition *WSP [ Comment ] NWLN ]
    ///  [ *WSP DesignerTag *WSP [ Comment ] NWLN ]
    ///  [ *WSP DesignerTypeInfoVer *WSP [ Comment ] NWLN ]
    ///  [ *WSP DesignerVisible *WSP [ Comment ] NWLN ]
    ///  [ *WSP DesignerWhatsThisButton *WSP [ Comment ] NWLN ]
    ///  [ *WSP DesignerWhatsThisHelp *WSP [ Comment ] NWLN ]
    /// Comment = "'" *ANYCHAR
    /// <Comment>: Specifies a user-readable comment.
    public struct DesignerProperties: ABNFGrammar {
        public let caption: (value: DesignerCaption, comment: String)?
        public let height: (value: DesignerHeight, comment: String)?
        public let left: (value: DesignerLeft, comment: String)?
        public let top: (value: DesignerTop, comment: String)?
        public let width: (value: DesignerWidth, comment: String)?
        public let enabled: (value: DesignerEnabled, comment: String)?
        public let helpContextId: (value: DesignerHelpContextId, comment: String)?
        public let rtl: (value: DesignerRTL, comment: String)?
        public let showModal: (value: DesignerShowModal, comment: String)?
        public let startupPosition: (value: DesignerStartupPosition, comment: String)?
        public let tag: (value: DesignerTag, comment: String)?
        public let typeInfoVer: (value: DesignerTypeInfoVer, comment: String)?
        public let visible: (value: DesignerVisible, comment: String)?
        public let whatsThisButton: (value: DesignerWhatsThisButton, comment: String)?
        public let whatsThisHelp: (value: DesignerWhatsThisHelp, comment: String)?
        
        init(stream: inout ABNFStream) throws {
            try WSP.readMultiple(stream: &stream)
            if stream.peek(string: "Caption") {
                self.caption = (try DesignerCaption(stream: &stream), try ANYCHAR.read(stream: &stream))
                try stream.require(NWLN.self)
            } else {
                self.caption = nil
            }
            
            try WSP.readMultiple(stream: &stream)
            if stream.peek(string: "ClientHeight") {
                self.height = (try DesignerHeight(stream: &stream), try ANYCHAR.read(stream: &stream))
                try stream.require(NWLN.self)
            } else {
                self.height = nil
            }
            
            try WSP.readMultiple(stream: &stream)
            if stream.peek(string: "ClientLeft") {
                self.left = (try DesignerLeft(stream: &stream), try ANYCHAR.read(stream: &stream))
                try stream.require(NWLN.self)
            } else {
                self.left = nil
            }
            
            try WSP.readMultiple(stream: &stream)
            if stream.peek(string: "ClientTop") {
                self.top = (try DesignerTop(stream: &stream), try ANYCHAR.read(stream: &stream))
                try stream.require(NWLN.self)
            } else {
                self.top = nil
            }
            
            try WSP.readMultiple(stream: &stream)
            if stream.peek(string: "ClientWidth") {
                self.width = (try DesignerWidth(stream: &stream), try ANYCHAR.read(stream: &stream))
                try stream.require(NWLN.self)
            } else {
                self.width = nil
            }
            
            try WSP.readMultiple(stream: &stream)
            if stream.peek(string: "Enabled") {
                self.enabled = (try DesignerEnabled(stream: &stream), try ANYCHAR.read(stream: &stream))
                try stream.require(NWLN.self)
            } else {
                self.enabled = nil
            }
            
            try WSP.readMultiple(stream: &stream)
            if stream.peek(string: "HelpContextID") {
                self.helpContextId = (try DesignerHelpContextId(stream: &stream), try ANYCHAR.read(stream: &stream))
                try stream.require(NWLN.self)
            } else {
                self.helpContextId = nil
            }
            
            try WSP.readMultiple(stream: &stream)
            if stream.peek(string: "RightToLeft") {
                self.rtl = (try DesignerRTL(stream: &stream), try ANYCHAR.read(stream: &stream))
                try stream.require(NWLN.self)
            } else {
                self.rtl = nil
            }
            
            try WSP.readMultiple(stream: &stream)
            if stream.peek(string: "ShowModal") {
                self.showModal = (try DesignerShowModal(stream: &stream), try ANYCHAR.read(stream: &stream))
                try stream.require(NWLN.self)
            } else {
                self.showModal = nil
            }
            
            try WSP.readMultiple(stream: &stream)
            if stream.peek(string: "StartUpPosition") {
                self.startupPosition = (try DesignerStartupPosition(stream: &stream), try ANYCHAR.read(stream: &stream))
                try stream.require(NWLN.self)
            } else {
                self.startupPosition = nil
            }
            
            try WSP.readMultiple(stream: &stream)
            if stream.peek(string: "Tag") {
                self.tag = (try DesignerTag(stream: &stream), try ANYCHAR.read(stream: &stream))
                try stream.require(NWLN.self)
            } else {
                self.tag = nil
            }
            
            try WSP.readMultiple(stream: &stream)
            if stream.peek(string: "TypeInfoVer") {
                self.typeInfoVer = (try DesignerTypeInfoVer(stream: &stream), try ANYCHAR.read(stream: &stream))
                try stream.require(NWLN.self)
            } else {
                self.typeInfoVer = nil
            }
            
            try WSP.readMultiple(stream: &stream)
            if stream.peek(string: "Visible") {
                self.visible = (try DesignerVisible(stream: &stream), try ANYCHAR.read(stream: &stream))
                try stream.require(NWLN.self)
            } else {
                self.visible = nil
            }
            
            try WSP.readMultiple(stream: &stream)
            if stream.peek(string: "WhatsThisButton") {
                self.whatsThisButton = (try DesignerWhatsThisButton(stream: &stream), try ANYCHAR.read(stream: &stream))
                try stream.require(NWLN.self)
            } else {
                self.whatsThisButton = nil
            }
            
            try WSP.readMultiple(stream: &stream)
            if stream.peek(string: "WhatsThisHelp") {
                self.whatsThisHelp = (try DesignerWhatsThisHelp(stream: &stream), try ANYCHAR.read(stream: &stream))
                try stream.require(NWLN.self)
            } else {
                self.whatsThisHelp = nil
            }
        }
    }
    
    /// [MS-OVBA] 2.3.5.2 DesignerCaption
    /// Specifies the title text of the designer.
    /// ABNF syntax:
    /// DesignerCaption = "Caption" EQ DQUOTE DesignerCaptionText DQUOTE
    /// DesignerCaptionText = *130QUOTEDCHAR
    public struct DesignerCaption: ABNFGrammar {
        public let designerCaptionText: String

        init(stream: inout ABNFStream) throws {
            try stream.require("Caption")
            try stream.require(EQ.self)
            try stream.require(DQUOTE.self)
            self.designerCaptionText = try QUOTEDCHAR.read(stream: &stream)
            guard self.designerCaptionText.count <= 130 else {
                throw VBAFileError.corrupted
            }
            try stream.require(DQUOTE.self)
        }
    }
    
    /// [MS-OVBA] 2.3.5.3 DesignerHeight
    /// Specifies the height of the designer in twips.
    /// ABNF syntax:
    /// DesignerHeight = "ClientHeight" EQ FLOAT
    public struct DesignerHeight: ABNFGrammar {
        public let value: FLOAT

        init(stream: inout ABNFStream) throws {
            try stream.require("ClientHeight")
            try stream.require(EQ.self)
            self.value = try stream.read()
        }
    }
    
    /// [MS-OVBA] 2.3.5.4 DesignerLeft
    /// Specifies the left edge of the designer in twips relative to the window specified by DesignerStartupPosition (section 2.3.5.11).
    /// ABNF syntax:
    /// DesignerLeft = "ClientLeft" EQ FLOAT
    public struct DesignerLeft: ABNFGrammar {
        public let value: FLOAT

        init(stream: inout ABNFStream) throws {
            try stream.require("ClientLeft")
            try stream.require(EQ.self)
            self.value = try stream.read()
        }
    }
    
    /// [MS-OVBA] 2.3.5.5 DesignerTop
    /// Specifies the position of the top edge of the designer in twips relative to the window specified by DesignerStartupPosition (section 2.3.5.11).
    /// ABNF syntax:
    /// DesignerTop = "ClientTop" EQ FLOAT
    public struct DesignerTop: ABNFGrammar {
        public let value: FLOAT

        init(stream: inout ABNFStream) throws {
            try stream.require("ClientTop")
            try stream.require(EQ.self)
            self.value = try stream.read()
        }
    }
    
    /// [MS-OVBA] 2.3.5.6 DesignerWidth
    /// Specifies the width of the designer in twips.
    /// ABNF Syntax:
    /// DesignerWidth = "ClientWidth" EQ FLOAT
    public struct DesignerWidth: ABNFGrammar {
        public let value: FLOAT

        init(stream: inout ABNFStream) throws {
            try stream.require("ClientWidth")
            try stream.require(EQ.self)
            self.value = try stream.read()
        }
    }
    
    /// [MS-OVBA] 2.3.5.7 DesignerEnabled
    /// Specifies whether the designer is enabled. The default is TRUE.
    /// ABNF syntax:
    /// DesignerEnabled = "Enabled" EQ VBABOOL
    public struct DesignerEnabled: ABNFGrammar {
        public let value: VBABOOL

        init(stream: inout ABNFStream) throws {
            try stream.require("Enabled")
            try stream.require(EQ.self)
            self.value = try stream.read()
        }
    }
    
    /// [MS-OVBA] 2.3.5.8 DesignerHelpContextId
    /// Specifies the Help topic identifier associated with this designer in the Help file as specified by ProjectHelpFile (section 2.3.1.9).
    /// ABNF syntax:
    /// DesignerHelpContextId = "HelpContextID" EQ INT32
    public struct DesignerHelpContextId: ABNFGrammar {
        public let value: INT32

        init(stream: inout ABNFStream) throws {
            try stream.require("HelpContextID")
            try stream.require(EQ.self)
            self.value = try stream.read()
        }
    }
    
    /// [MS-OVBA] 2.3.5.9 DesignerRTL
    /// Specifies that the designer be shown with right and left coordinates reversed for right-to-left language use.
    /// ABNF syntax:
    /// DesignerRTL = "RightToLeft" EQ VBABOOL
    public struct DesignerRTL: ABNFGrammar {
        public let value: VBABOOL

        init(stream: inout ABNFStream) throws {
            try stream.require("RightToLeft")
            try stream.require(EQ.self)
            self.value = try stream.read()
        }
    }
    
    /// [MS-OVBA] 2.3.5.10 DesignerShowModal
    /// Specifies whether the designer is a modal window. The default is TRUE.
    /// ABNF syntax:
    /// DesignerShowModal = "ShowModal" EQ VBABOOL
    public struct DesignerShowModal: ABNFGrammar {
        public let value: VBABOOL

        init(stream: inout ABNFStream) throws {
            try stream.require("ShowModal")
            try stream.require(EQ.self)
            self.value = try stream.read()
        }
    }
    
    /// [MS-OVBA] 2.3.5.11 DesignerStartupPosition
    /// Specifies the startup position of the designer as follows.
    /// ABNF syntax:
    /// DesignerStartupPosition = "StartUpPosition" EQ RelativeParent
    /// RelativeParent = "0" / "1" / "2" / "3"
    /// <RelativeParent>: Specifies the window used to determine the relative starting coordinates of the control window.
    /// MUST be one of the following values:
    /// Value Meaning
    /// "0" "Manual" mode. DesignerTop (section 2.3.5.5) and DesignerLeft (section 2.3.5.4) coordinates of the designer are relative to the desktop
    /// window.
    /// "1" "CenterOwner" mode. Center the designer relative to its parent window.
    /// "2" "Center" mode. Center the designer relative to the desktop window.
    /// "3" "WindowsDefault" mode. Place the designer in the upper-left corner of screen.
    public struct DesignerStartupPosition: ABNFGrammar {
        public let relativeParent: RelativeParent

        init(stream: inout ABNFStream) throws {
            try stream.require("StartUpPosition")
            try stream.require(EQ.self)
            guard let relativeParent = RelativeParent(rawValue: try stream.readString(count: 1)) else {
                throw VBAFileError.corrupted
            }
            
            self.relativeParent = relativeParent
        }
        
        /// <RelativeParent>: Specifies the window used to determine the relative starting coordinates of the control window.
        /// MUST be one of the following values:
        public enum RelativeParent: String {
            /// "0" "Manual" mode. DesignerTop (section 2.3.5.5) and DesignerLeft (section 2.3.5.4) coordinates of the designer are relative to the desktop
            /// window.
            case manual = "0"
            
            /// "1" "CenterOwner" mode. Center the designer relative to its parent window.
            case centerOwner = "1"
            
            /// "2" "Center" mode. Center the designer relative to the desktop window.
            case center = "2"
            
            /// "3" "WindowsDefault" mode. Place the designer in the upper-left corner of screen.
            case windowsDefault = "3"
        }
    }

    /// [MS-OVBA] 2.3.5.12 DesignerTag
    /// Specifies user-defined data associated with the designer.
    /// ABNF syntax:
    /// DesignerTag = "Tag" EQ DQUOTE DesignerTagText DQUOTE
    /// DesignerTagText = *130QUOTEDCHAR
    public struct DesignerTag: ABNFGrammar {
        public let designerTagText: String

        init(stream: inout ABNFStream) throws {
            try stream.require("Tag")
            try stream.require(EQ.self)
            try stream.require(DQUOTE.self)
            self.designerTagText = try QUOTEDCHAR.read(stream: &stream)
            guard self.designerTagText.count <= 130 else {
                throw VBAFileError.corrupted
            }
            try stream.require(DQUOTE.self)
        }
    }
    
    /// [MS-OVBA] 2.3.5.13 DesignerTypeInfoVer
    /// Specifies the number of times the designer has been changed and saved. The default is 0.
    /// ABNF syntax:
    /// DesignerTypeInfoVer = "TypeInfoVer" EQ INT32
    public struct DesignerTypeInfoVer: ABNFGrammar {
        public let value: INT32

        init(stream: inout ABNFStream) throws {
            try stream.require("TypeInfoVer")
            try stream.require(EQ.self)
            self.value = try stream.read()
        }
    }
    
    /// [MS-OVBA] 2.3.5.14 DesignerVisible
    /// Specifies whether the designer is visible. The default is TRUE.
    /// ABNF syntax:
    /// DesignerVisible = "Visible" EQ VBABOOL
    public struct DesignerVisible: ABNFGrammar {
        public let value: VBABOOL

        init(stream: inout ABNFStream) throws {
            try stream.require("Visible")
            try stream.require(EQ.self)
            self.value = try stream.read()
        }
    }
    
    /// [MS-OVBA] 2.3.5.15 DesignerWhatsThisButton
    /// Specifies whether a help button is shown for the designer. The default is FALSE.
    /// ABNF syntax:
    /// DesignerWhatsThisButton = "WhatsThisButton" EQ VBABOOL
    public struct DesignerWhatsThisButton: ABNFGrammar {
        public let value: VBABOOL

        init(stream: inout ABNFStream) throws {
            try stream.require("WhatsThisButton")
            try stream.require(EQ.self)
            self.value = try stream.read()
        }
    }
    
    /// [MS-OVBA] 2.3.5.16 DesignerWhatsThisHelp
    /// Specifies whether a help topic is associated with this designer. The Help topic identifier is specified by DesignerHelpContextId
    /// (section 2.3.5.8).
    /// ABNF syntax:
    /// DesignerWhatsThisHelp = "WhatsThisHelp" EQ VBABOOL
    public struct DesignerWhatsThisHelp: ABNFGrammar {
        public let value: VBABOOL

        init(stream: inout ABNFStream) throws {
            try stream.require("WhatsThisHelp")
            try stream.require(EQ.self)
            self.value = try stream.read()
        }
    }
}
