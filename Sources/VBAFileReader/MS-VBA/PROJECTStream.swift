//
//  PROJECTStream.swift
//  
//
//  Created by Hugh Bellamy on 26/11/2020.
//

import DataStream
import WindowsDataTypes

/// [MS-OVBA] 2.3.1 PROJECT Stream: Project Information
/// The PROJECT stream specifies properties of the VBA project.
/// This stream is an array of bytes that specifies properties of the VBA project. MUST contain MBCS characters encoded using the code page
/// specified in PROJECTCODEPAGE (section 2.3.4.2.1.4).
/// ABNF syntax:
/// VBAPROJECTText = ProjectProperties NWLN
/// HostExtenders
/// [NWLN ProjectWorkspace]
public struct PROJECTStream {
    public let projectProperties: ProjectProperties
    public let hostExtenders: HostExtenders
    public let projectWorkspace: ProjectWorkspace?
    
    public init(dataStream: inout DataStream, count: Int) throws {
        let startPosition = dataStream.position

        var stream = ABNFStream(dataStream: &dataStream)
        
        self.projectProperties = try stream.read(ProjectProperties.self)
        try stream.require(NWLN.self)
        self.hostExtenders = try stream.read(HostExtenders.self)
        
        if stream.position - startPosition == count {
            self.projectWorkspace = nil
            return
        }
        
        try stream.require(NWLN.self)
        self.projectWorkspace = try stream.read(ProjectWorkspace.self)
        
        guard stream.position - startPosition == count else {
            throw VBAFileError.corrupted
        }
    }

    /// [MS-OVBA] 2.3.1.1 ProjectProperties
    /// Specifies project-wide properties.
    /// ABNF syntax:
    /// ProjectProperties = ProjectId
    ///  *ProjectItem
    /// [ProjectHelpFile]
    ///  [ProjectExeName32]
    ///  ProjectName
    /// ProjectHelpId
    ///  [ProjectDescription]
    ///  [ProjectVersionCompat32]
    ///  ProjectProtectionState
    /// ProjectPassword
    /// ProjectVisibilityState
    public struct ProjectProperties: ABNFGrammar {
        public let projectId: ProjectId
        public let projectItems: [ProjectItem]
        public let projectHelpFile: ProjectHelpFile?
        public let projectExeName32: ProjectExeName32?
        public let projectName: ProjectName
        public let projectHelpId: ProjectHelpId
        public let projectDescription: ProjectDescription?
        public let projectVersionCompat32: ProjectVersionCompat32?
        public let projectProtectionState: ProjectProtectionState
        public let projectPassword: ProjectPassword
        public let projectVisibilityState: ProjectVisibilityState
        
        init(stream: inout ABNFStream) throws {
            self.projectId = try stream.read(ProjectId.self)
            
            var projectItems: [ProjectItem] = []
            while true {
                let item = try stream.read(ProjectItem.self)
                if case .none = item {
                    break
                }

                projectItems.append(item)
            }
            
            self.projectItems = projectItems
            
            if stream.peek(string: "HelpFile=") {
                self.projectHelpFile = try ProjectHelpFile(stream: &stream)
            } else {
                self.projectHelpFile = nil
            }
            
            if stream.peek(string: "ExeName32=") {
                self.projectExeName32 = try ProjectExeName32(stream: &stream)
            } else {
                self.projectExeName32 = nil
            }
            
            self.projectName = try ProjectName(stream: &stream)

            self.projectHelpId = try ProjectHelpId(stream: &stream)
            
            if stream.peek(string: "Description=") {
                self.projectDescription = try ProjectDescription(stream: &stream)
            } else {
                self.projectDescription = nil
            }
            
            if stream.peek(string: "VersionCompatible32=") {
                self.projectVersionCompat32 = try ProjectVersionCompat32(stream: &stream)
            } else {
                self.projectVersionCompat32 = nil
            }
            
            self.projectProtectionState = try ProjectProtectionState(stream: &stream)
            
            self.projectPassword = try ProjectPassword(stream: &stream)
            
            self.projectVisibilityState = try ProjectVisibilityState(stream: &stream)
        }
        
        /// ProjectItem = ( ProjectModule /
        ///  ProjectPackage ) NWLN
        public enum ProjectItem: ABNFGrammar {
            case module(_: ProjectModule)
            case package(_: ProjectPackage)
            case none
            
            init(stream: inout ABNFStream) throws {
                if stream.peek(string: "Document=") ||
                    stream.peek(string: "Module=") ||
                    stream.peek(string: "Class=") ||
                    stream.peek(string: "BaseClass=") {
                    self = .module(try ProjectModule(stream: &stream))
                    try stream.require(NWLN.self)
                } else if stream.peek(string: "Package=") {
                    self = .package(try ProjectPackage(stream: &stream))
                    try stream.require(NWLN.self)
                } else {
                    self = .none
                }
            }
        }
    }
    
    /// [MS-OVBA] 2.3.1.2 ProjectId
    /// Specifies the class identifier (CLSID) for the VBA project.
    /// ABNF syntax:
    /// ProjectId = "ID=" DQUOTE ProjectCLSID DQUOTE NWLN
    /// ProjectCLSID = GUID
    /// <ProjectCLSID>: Specifies the class identifier (CLSID) of the VBA project’s Automation type library. MUST be
    /// "{00000000-0000-0000-0000-000000000000}" when ProjectPassword (section 2.3.1.16) specifies a password hash.
    public struct ProjectId: ABNFGrammar {
        public let projectCLSID: GUID
        
        init(stream: inout ABNFStream) throws {
            try stream.require("ID=")
            try stream.require(DQUOTE.self)
            
            self.projectCLSID = try stream.read()
            
            try stream.require(DQUOTE.self)
            try stream.require(NWLN.self)
        }
    }
    
    /// [MS-OVBA] 2.3.1.3 ProjectModule
    /// Specifies a module that contains VBA language source code as specified in [MS-VBAL] section 4.2.
    /// ABNF syntax:
    /// ProjectModule = ( ProjectDocModule /
    ///  ProjectStdModule /
    ///  ProjectClassModule /
    ///  ProjectDesignerModule )
    /// <ProjectModule>: Specifies the name and type of a specific module. MUST have a corresponding
    /// MODULE Record (section 2.3.4.2.3.2) in the dir Stream (section 2.3.4.2).
    public enum ProjectModule: ABNFGrammar {
        case none
        case docModule(_: ProjectDocModule)
        case stdModule(_: ProjectStdModule)
        case classModule(_: ProjectClassModule)
        case designerModule(_: ProjectDesignerModule)
        
        init(stream: inout ABNFStream) throws {
            if stream.peek(string: "Document=") {
                self = .docModule(_: try ProjectDocModule(stream: &stream))
            } else if stream.peek(string: "Module=") {
                self = .stdModule(_: try ProjectStdModule(stream: &stream))
            } else if stream.peek(string: "Class=") {
                self = .classModule(_: try ProjectClassModule(stream: &stream))
            } else if stream.peek(string: "BaseClass=") {
                self = .designerModule(_: try ProjectDesignerModule(stream: &stream))
            } else {
                self = .none
            }
        }
    }
    
    /// [MS-OVBA] 2.3.1.4 ProjectDocModule
    /// Specifies a module that extends a document module.
    /// ABNF syntax:
    /// ProjectDocModule = "Document=" ModuleIdentifier %x2f DocTlibVer
    /// DocTlibVer = HEXINT32
    /// <DocTlibVer>: Specifies the document module’s Automation server version as specified by [MSOAUT].
    public struct ProjectDocModule: ABNFGrammar {
        public let moduleIdentifier: ModuleIdentifier
        public let docTlibVer: HEXINT32
        
        init(stream: inout ABNFStream) throws {
            try stream.require("Document=")
            self.moduleIdentifier = try stream.read(ModuleIdentifier.self)
            try stream.require("/")
            self.docTlibVer = try stream.read()
        }
    }

    /// [MS-OVBA] 2.3.1.5 ProjectStdModule
    /// Specifies a procedural module.
    /// ABNF syntax:
    /// ProjectStdModule = "Module=" ModuleIdentifier
    public struct ProjectStdModule: ABNFGrammar {
        public let moduleIdentifier: ModuleIdentifier
        
        init(stream: inout ABNFStream) throws {
            try stream.require("Module=")
            self.moduleIdentifier = try stream.read()
        }
    }
    
    /// [MS-OVBA] 2.3.1.6 ProjectClassModule
    /// Specifies a class module.
    /// ABNF syntax:
    /// ProjectClassModule = "Class=" ModuleIdentifier
    public struct ProjectClassModule: ABNFGrammar {
        public let moduleIdentifier: ModuleIdentifier
        
        init(stream: inout ABNFStream) throws {
            try stream.require("Class=")
            self.moduleIdentifier = try stream.read()
        }
    }
    
    /// [MS-OVBA] 2.3.1.7 ProjectDesignerModule
    /// Specifies a designer module.
    /// ABNF syntax:
    /// ProjectDesignerModule = "BaseClass=" ModuleIdentifier
    public struct ProjectDesignerModule: ABNFGrammar {
        public let moduleIdentifier: ModuleIdentifier
        
        init(stream: inout ABNFStream) throws {
            try stream.require("BaseClass=")
            self.moduleIdentifier = try stream.read()
        }
    }
    
    /// [MS-OVBA] 2.3.1.8 ProjectPackage
    /// Specifies the class identifier (CLSID) for a designer extended by one or more modules.
    /// ABNF syntax:
    /// ProjectPackage = "Package=" GUID
    public struct ProjectPackage: ABNFGrammar {
        public let clsid: GUID
        
        init(stream: inout ABNFStream) throws {
            try stream.require("Package=")
            self.clsid = try stream.read()
        }
    }

    /// [MS-OVBA] 2.3.1.9 ProjectHelpFile
    /// Specifies a path to a Help file associated with this VBA project. MUST be the same value as specified in PROJECTHELPFILEPATH
    /// (section 2.3.4.2.1.7). MUST be present if PROJECTHELPFILEPATH specifies a value.
    /// ABNF syntax:
    /// ProjectHelpFile = "HelpFile=" PATH NWLN
    public struct ProjectHelpFile: ABNFGrammar {
        public let path: PATH
        
        init(stream: inout ABNFStream) throws {
            try stream.require("HelpFile=")
            self.path = try stream.read()
            try stream.require(NWLN.self)
        }
    }
    
    /// [MS-OVBA] 2.3.1.10 ProjectExeName32
    /// Specifies a path. MUST be ignored.
    /// ABNF syntax:
    /// ProjectExeName32 = "ExeName32=" PATH NWLN
    public struct ProjectExeName32: ABNFGrammar {
        public let path: PATH
        
        init(stream: inout ABNFStream) throws {
            try stream.require("ExeName32=")
            self.path = try stream.read()
            try stream.require(NWLN.self)
        }
    }
    
    /// [MS-OVBA] 2.3.1.11 ProjectName
    /// Specifies the short name of the VBA project.
    /// ABNF syntax:
    /// ProjectName = "Name=" DQUOTE ProjectIdentifier DQUOTE NWLN
    /// ProjectIdentifier = 1*128QUOTEDCHAR
    /// <ProjectIdentifier>: Specifies the name of the VBA project. MUST be less than or equal to 128 characters long. MUST be the
    /// same value as specified in PROJECTNAME (section 2.3.4.2.1.5). SHOULD be an identifier as specified by [MS-VBAL] section 3.3.5.
    /// MAY<3> be any string of characters.
    public struct ProjectName: ABNFGrammar {
        public let projectIdentifier: String
        
        init(stream: inout ABNFStream) throws {
            try stream.require("Name=")
            try stream.require(DQUOTE.self)
            self.projectIdentifier = try QUOTEDCHAR.read(stream: &stream)
            guard self.projectIdentifier.count <= 128 else {
                throw VBAFileError.corrupted
            }

            try stream.require(DQUOTE.self)
            try stream.require(NWLN.self)
        }
    }
    
    /// [MS-OVBA] 2.3.1.12 ProjectHelpId
    /// Specifies a Help topic identifier in ProjectHelpFile (section 2.3.1.9) associated with this VBA project.
    /// ABNF syntax:
    /// ProjectHelpId = "HelpContextID=" DQUOTE TopicId DQUOTE NWLN
    /// TopicId = INT32
    /// <TopicId>: Specifies a Help topic identifier. MUST be the same value as specified in PROJECTHELPCONTEXT (section 2.3.4.2.1.8).
    public struct ProjectHelpId: ABNFGrammar {
        public let topicId: INT32
        
        init(stream: inout ABNFStream) throws {
            try stream.require("HelpContextID=")
            try stream.require(DQUOTE.self)
            self.topicId = try stream.read()
            try stream.require(DQUOTE.self)
            try stream.require(NWLN.self)
        }
    }
    
    /// [MS-OVBA] 2.3.1.13 ProjectDescription
    /// Specifies the description of the VBA project.
    /// ABNF syntax:
    /// ProjectDescription = "Description=" DQUOTE DescriptionText DQUOTE NWLN
    /// DescriptionText = *2000QUOTEDCHAR
    /// <DescriptionText>: MUST be the same value as specified in PROJECTDOCSTRING (section 2.3.4.2.1.6).
    public struct ProjectDescription: ABNFGrammar {
        public let descriptionText: String
        
        init(stream: inout ABNFStream) throws {
            try stream.require("Description=")
            try stream.require(DQUOTE.self)
            self.descriptionText = try QUOTEDCHAR.read(stream: &stream)
            guard self.descriptionText.count <= 200 else {
                throw VBAFileError.corrupted
            }

            try stream.require(DQUOTE.self)
            try stream.require(NWLN.self)
        }
    }
    
    /// [MS-OVBA] 2.3.1.14 ProjectVersionCompat32
    /// Specifies the storage format version of the VBA project. MAY be missing<4>.
    /// ABNF syntax:
    /// ProjectVersionCompat32 = "VersionCompatible32=" DQUOTE "393222000" DQUOTE NWLN
    public struct ProjectVersionCompat32: ABNFGrammar {
        init(stream: inout ABNFStream) throws {
            try stream.require("VersionCompatible32=")
            try stream.require(DQUOTE.self)
            try stream.require("393222000")
            try stream.require(DQUOTE.self)
            try stream.require(NWLN.self)
        }
    }
    
    /// [MS-OVBA] 2.3.1.15 ProjectProtectionState
    /// Specifies whether access to the VBA project was restricted by the user, the VBA host application, or the VBA project editor.
    /// ABNF syntax:
    /// ProjectProtectionState = "CMG=" DQUOTE EncryptedState DQUOTE NWLN
    /// EncryptedState = 22*28HEXDIG
    /// <EncryptedState>: Specifies whether access to the VBA project was restricted by the user, the VBA host application, or the VBA
    /// project editor, obfuscated by Data Encryption (section 2.4.3.2).
    /// The Data parameter for Data Encryption (section 2.4.3.2) SHOULD be four bytes that specify the protection state of the VBA project.
    /// MAY<5> be 0x00000000. The Length parameter for Data Encryption (section 2.4.3.2) MUST be 4.
    /// <DescriptionText>: MUST be the same value as specified in PROJECTDOCSTRING (section 2.3.4.2.1.6).
    /// Values for Data are defined by the following bits:
    public struct ProjectProtectionState: ABNFGrammar {
        public let encryptedState: String
        
        init(stream: inout ABNFStream) throws {
            try stream.require("CMG=")
            try stream.require(DQUOTE.self)
            self.encryptedState = try QUOTEDCHAR.read(stream: &stream)
            try stream.require(DQUOTE.self)
            try stream.require(NWLN.self)
        }
    }
    
    /// [MS-OVBA] 2.3.1.16 ProjectPassword
    /// Specifies the password hash of the VBA project.
    /// The syntax of ProjectPassword is defined as follows.
    /// ProjectPassword = "DPB=" DQUOTE EncryptedPassword DQUOTE NWLN
    /// EncryptedPassword = 16*HEXDIG
    /// <EncryptedPassword>: Specifies the password protection for the VBA project.
    /// A VBA project without a password MUST use 0x00 for the Data parameter for Data Encryption (section 2.4.3.2) and the Length
    /// parameter MUST be 1.
    /// A VBA project with a password SHOULD specify the password hash of the VBA project, obfuscated by Data Encryption (section 2.4.3.2).
    /// The Data parameter for Data Encryption (section 2.4.3.2) MUST be an array of bytes that specifies a Hash Data Structure (section 2.4.4.1)
    /// and the Length parameter for Data Encryption MUST be 29. The Hash Data Structure (section 2.4.4.1) specifies a hash key and
    /// password hash encoded to remove null bytes as specified by section 2.4.4.
    /// A VBA project with a password MAY<6> specify the plain text password of the VBA project, obfuscated by Data Encryption (section
    /// 2.4.3.2). In this case, the Data parameter Data Encryption (section 2.4.3.2) MUST be an array of bytes that specifies a null terminated
    /// password string encoded using MBCS using the code page specified by PROJECTCODEPAGE (section 2.3.4.2.1.4), and a Length
    /// parameter equal to the number of bytes in the password string including the terminating null character.
    /// When the data specified by <EncryptpedPassword> is a password hash, ProjectId.ProjectCLSID (section 2.3.1.2) MUST be
    /// "{00000000-0000-0000-0000-000000000000}".
    public struct ProjectPassword: ABNFGrammar {
        public let encryptedPassword: String
        
        init(stream: inout ABNFStream) throws {
            try stream.require("DPB=")
            try stream.require(DQUOTE.self)
            self.encryptedPassword = try QUOTEDCHAR.read(stream: &stream)
            try stream.require(DQUOTE.self)
            try stream.require(NWLN.self)
        }
    }
    
    /// [MS-OVBA] 2.3.1.17 ProjectVisibilityState
    /// Specifies whether the VBA project is visible.
    /// ABNF syntax:
    /// ProjectVisibilityState = "GC=" DQUOTE EncryptedProjectVisibility DQUOTE NWLN
    /// EncryptedProjectVisibility = 16*22HEXDIG
    /// <EncryptedProjectVisibility>: Specifies whether the VBA project is visible, obfuscated by Data Encryption (section 2.4.3.2).
    /// The Data parameter for Data Encryption (section 2.4.3.2) is one byte that specifies the visibility state of the VBA project. The Length
    /// parameter for Data Encryption (section 2.4.3.2) MUST be 1.
    /// Values for Data are:
    /// Value Meaning
    /// 0x00 VBA project is NOT visible. <ProjectProtectionState>.fVBEProtected (section 2.3.1.15) MUST be TRUE.
    /// 0xFF VBA project is visible.
    /// The default is 0xFF.
    public struct ProjectVisibilityState: ABNFGrammar {
        public let encryptedProjectVisibility: String
        
        init(stream: inout ABNFStream) throws {
            try stream.require("GC=")
            try stream.require(DQUOTE.self)
            self.encryptedProjectVisibility = try QUOTEDCHAR.read(stream: &stream)
            try stream.require(DQUOTE.self)
            try stream.require(NWLN.self)
        }
    }
    
    /// [MS-OVBA] 2.3.1.18 HostExtenders
    /// Specifies a list of host extenders.
    /// ABNF syntax:
    /// HostExtenders = "[Host Extender Info]" NWLN
    ///  *HostExtenderRef
    /// HostExtenderRef = ExtenderIndex "=" ExtenderGuid ";"
    ///  LibName ";" CreationFlags NWLN
    /// ExtenderIndex = HEXINT32
    /// ExtenderGuid = GUID
    /// LibName = "VBE" / *(%x21-3A / %x3C-FF)
    /// CreationFlags = HEXINT32
    /// <HostExtenderRef>: Specifies a reference to an aggregatable server’s Automation type library.
    /// <ExtenderIndex>: Specifies the index of the host extender entry. MUST be unique to the list of HostExtenders.
    /// <ExtenderGuid>: Specifies the GUID of the Automation type library to extend.
    /// <LibName>: Specifies a host-provided Automation type library name. "VBE" specifies a built in name for the VBA Automation type library.
    /// <CreationFlags>: Specifies a host-provided flag as follows:
    /// Value Meaning
    /// 0x00000000
    /// MUST NOT create a new extended type library for the aggregatable server if one is
    /// already available to the VBA environment.
    /// 0x00000001 MUST create a new extended type library for the aggregatable server.
    public struct HostExtenders: ABNFGrammar {
        public let extenders: [HostExtenderRef]
        
        init(stream: inout ABNFStream) throws {
            try stream.require("[Host Extender Info]")
            try stream.require(NWLN.self)
            
            var extenders: [HostExtenderRef] = []
            while stream.position < stream.count {
                if stream.position + 2 < stream.count {
                    let nextString = try stream.peekString(count: 2)
                    if nextString == "\r\n" || nextString == "\n\r" {
                        break
                    }
                }

                extenders.append(try stream.read())
            }
            
            self.extenders = extenders
        }
        
        /// HostExtenderRef = ExtenderIndex "=" ExtenderGuid ";" LibName ";" CreationFlags NWLN
        /// ExtenderIndex = HEXINT32
        /// ExtenderGuid = GUID
        /// LibName = "VBE" / *(%x21-3A / %x3C-FF)
        /// CreationFlags = HEXINT32
        public struct HostExtenderRef: ABNFGrammar {
            public let index: HEXINT32
            public let guid: GUID
            public let libName: String
            public let creationFlags: HEXINT32
            
            init(stream: inout ABNFStream) throws {
                self.index = try stream.read()
                try stream.require("=")
                self.guid = try stream.read()
                try stream.require(";")
                guard let index = stream.nextIndex(string: ";") else {
                    throw VBAFileError.corrupted
                }
                
                self.libName = try stream.readString(count: index - stream.position)
                try stream.require(";")
                self.creationFlags = try stream.read()
                try stream.require(NWLN.self)
            }
        }
    }
    
    /// [MS-OVBA] 2.3.1.19 ProjectWorkspace
    /// Specifies a list of module editor window states.
    /// ABNF syntax:
    /// ProjectWorkspace = "[Workspace]" NWLN *ProjectWindowRecord
    public struct ProjectWorkspace: ABNFGrammar {
        public let windowRecords: [ProjectWindowRecord]
        
        init(stream: inout ABNFStream) throws {
            try stream.require("[Workspace]")
            try stream.require(NWLN.self)
            
            var windowRecords: [ProjectWindowRecord] = []
            while stream.position < stream.count {
                if stream.position + 2 < stream.count {
                    let nextString = try stream.peekString(count: 2)
                    if nextString == "\r\n" || nextString == "\n\r" {
                        break
                    }
                }

                windowRecords.append(try stream.read())
            }
            
            self.windowRecords = windowRecords
        }
    }
    
    /// [MS-OVBA] 2.3.1.20 ProjectWindowRecord
    /// Specifies the coordinates and state of a module editor window.
    /// ABNF syntax:
    /// ProjectWindowRecord = ModuleIdentifier "=" ProjectWindowState NWLN
    /// ProjectWindowState = CodeWindow [ ", " DesignerWindow ]
    /// CodeWindow = ProjectWindow
    /// DesignerWindow = ProjectWindow
    /// ProjectWindow = WindowLeft ", "
    ///  WindowTop ", "
    ///  WindowRight ", "
    ///  WindowBottom ", "
    ///  WindowState
    /// WindowLeft = INT32
    /// WindowTop = INT32
    /// WindowRight = INT32
    /// WindowBottom = INT32
    /// WindowState = ["C"] ["Z"] ["I"]
    /// <ModuleIdentifier>: Specifies the name of the module. MUST have a corresponding ProjectModule (section 2.3.1.3).
    /// <CodeWindow>: Specifies the coordinates and the state of a window used to edit the source code of a module.
    /// <DesignerWindow>: Specifies the coordinates and the state of a window used to edit the designer associated with a module.
    /// <WindowLeft>: Specifies the distance of the left edge of a window relative to a parent window.
    /// <WindowTop>: Specifies the distance of the top edge of a window relative to a parent window.
    /// <WindowRight>: Specifies the distance of the right edge of a window relative to a parent window.
    /// <WindowBottom>: Specifies the distance of the bottom edge of a window relative to a parent window.
    /// <WindowState>: Specifies the window state. Values are defined as follows:
    public struct ProjectWindowRecord: ABNFGrammar {
        public let moduleIdentifier: ModuleIdentifier
        public let codeWindow: ProjectWindow
        public let designerWindow: ProjectWindow?

        init(stream: inout ABNFStream) throws {
            self.moduleIdentifier = try stream.read()
            try stream.require("=")
            self.codeWindow = try stream.read()
            if stream.peek(string: ", ") {
                try stream.require(", ")
                self.designerWindow = try stream.read()
            } else {
                self.designerWindow = nil
            }
            
            try stream.require(NWLN.self)
        }
        
        public struct ProjectWindow: ABNFGrammar {
            public let windowLeft: INT32
            public let windowTop: INT32
            public let windowRight: INT32
            public let windowBottom: INT32
            public let windowState: State

            init(stream: inout ABNFStream) throws {
                self.windowLeft = try stream.read()
                try stream.require(", ")
                self.windowTop = try stream.read()
                try stream.require(", ")
                self.windowRight = try stream.read()
                try stream.require(", ")
                self.windowBottom = try stream.read()
                try stream.require(", ")
                
                guard let windowState = State(rawValue: try stream.readString(count: 1)) else {
                    throw VBAFileError.corrupted
                }
                
                self.windowState = windowState
            }
            
            public enum State: String {
                /// C Closed.
                case closed = "C"
                
                /// Z Zoomed to fill the available viewing area.
                case zoomed = "Z"
                
                /// I Minimized to an icon.
                case minimized = "I"
            }
        }
    }
}
