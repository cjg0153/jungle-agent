# JUNGLE_SMART_REFLOW_MUTATION_BOUNDARY_V1
# JUNGLE_SMART_REFLOW_MUTATION_BOUNDARY_V2
#
# V2:
# - stable slide IDs
# - semantic paragraph text
# - source shape identity preservation
# - exact group membership
# - dual-tagged decorative additions only
#
# Pure OpenXML / ZIP validation.
# No PowerPoint COM.

$ErrorActionPreference =
    "Stop"


try {
    Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
}
catch {}


try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
}
catch {}


function Get-JungleSha256 {

    param(
        [Parameter(Mandatory=$true)]
        [string]$Path
    )

    return (
        Get-FileHash `
            -LiteralPath $Path `
            -Algorithm SHA256
    ).Hash
}


function Get-JungleStringSha256 {

    param(
        [Parameter(Mandatory=$true)]
        [AllowEmptyString()]
        [string]$Text
    )

    $encoding =
        New-Object System.Text.UTF8Encoding(
            $false
        )

    $bytes =
        $encoding.GetBytes(
            $Text
        )

    $sha =
        [System.Security.Cryptography.SHA256]::Create()

    try {

        $hash =
            $sha.ComputeHash(
                $bytes
            )
    }
    finally {

        $sha.Dispose()
    }

    return (
        -join (
            $hash |
            ForEach-Object {
                $_.ToString("x2")
            }
        )
    ).ToUpperInvariant()
}


function Get-JungleZipEntryText {

    param(
        [Parameter(Mandatory=$true)]
        [object]$Zip,

        [Parameter(Mandatory=$true)]
        [string]$EntryPath
    )

    $entry =
        $Zip.GetEntry(
            $EntryPath
        )

    if ($null -eq $entry) {

        throw (
            "PPTX ZIP entry missing: " +
            $EntryPath
        )
    }

    $stream =
        $entry.Open()

    $reader =
        New-Object System.IO.StreamReader(
            $stream,
            [System.Text.Encoding]::UTF8,
            $true
        )

    try {

        return $reader.ReadToEnd()
    }
    finally {

        $reader.Dispose()
        $stream.Dispose()
    }
}


function Resolve-JungleZipRelativePath {

    param(
        [Parameter(Mandatory=$true)]
        [string]$BaseDirectory,

        [Parameter(Mandatory=$true)]
        [string]$Target
    )

    $cleanTarget =
        $Target.Replace(
            "\",
            "/"
        )

    if ($cleanTarget.StartsWith("/")) {

        $combined =
            $cleanTarget.TrimStart(
                "/"
            )
    }
    else {

        $combined =
            (
                $BaseDirectory.TrimEnd("/") +
                "/" +
                $cleanTarget
            )
    }

    $parts =
        $combined.Split(
            [char]"/"
        )

    $stack =
        @()

    foreach ($part in $parts) {

        if (
            [string]::IsNullOrWhiteSpace(
                $part
            ) `
            -or
            $part -eq "."
        ) {

            continue
        }

        if ($part -eq "..") {

            if ($stack.Count -eq 0) {

                throw (
                    "Invalid ZIP relative path: " +
                    $Target
                )
            }

            if ($stack.Count -eq 1) {

                $stack =
                    @()
            }
            else {

                $stack =
                    @(
                        $stack[
                            0..($stack.Count - 2)
                        ]
                    )
            }

            continue
        }

        $stack +=
            $part
    }

    return (
        $stack -join "/"
    )
}


function Get-JungleCnvPrNode {

    param(
        [Parameter(Mandatory=$true)]
        [System.Xml.XmlNode]$ShapeNode,

        [Parameter(Mandatory=$true)]
        [System.Xml.XmlNamespaceManager]$Ns
    )

    switch (
        [string]$ShapeNode.LocalName
    ) {

        "sp" {

            return $ShapeNode.SelectSingleNode(
                "./p:nvSpPr/p:cNvPr",
                $Ns
            )
        }

        "pic" {

            return $ShapeNode.SelectSingleNode(
                "./p:nvPicPr/p:cNvPr",
                $Ns
            )
        }

        "graphicFrame" {

            return $ShapeNode.SelectSingleNode(
                "./p:nvGraphicFramePr/p:cNvPr",
                $Ns
            )
        }

        "cxnSp" {

            return $ShapeNode.SelectSingleNode(
                "./p:nvCxnSpPr/p:cNvPr",
                $Ns
            )
        }

        "grpSp" {

            return $ShapeNode.SelectSingleNode(
                "./p:nvGrpSpPr/p:cNvPr",
                $Ns
            )
        }
    }

    return $null
}


function Get-JungleReflowSemanticTextInfo {

    param(
        [Parameter(Mandatory=$true)]
        [System.Xml.XmlDocument]$SlideXml
    )

    $ns =
        New-Object System.Xml.XmlNamespaceManager(
            $SlideXml.NameTable
        )

    $ns.AddNamespace(
        "a",
        "http://schemas.openxmlformats.org/drawingml/2006/main"
    )

    $paragraphNodes =
        $SlideXml.SelectNodes(
            "//a:p",
            $ns
        )

    $paragraphTexts =
        @()

    foreach ($paragraph in $paragraphNodes) {

        $textNodes =
            $paragraph.SelectNodes(
                ".//a:t",
                $ns
            )

        $paragraphText =
            ""

        foreach ($textNode in $textNodes) {

            $paragraphText +=
                [string]$textNode.InnerText
        }

        $paragraphTexts +=
            $paragraphText
    }

    $semanticText =
        $paragraphTexts -join [char]0x001E

    return [pscustomobject]@{

        paragraph_count =
            $paragraphTexts.Count

        semantic_text =
            $semanticText

        semantic_text_hash =
            Get-JungleStringSha256 `
                -Text $semanticText
    }
}


function Get-JungleReflowShapeInfo {

    param(
        [Parameter(Mandatory=$true)]
        [System.Xml.XmlDocument]$SlideXml
    )

    $ns =
        New-Object System.Xml.XmlNamespaceManager(
            $SlideXml.NameTable
        )

    $ns.AddNamespace(
        "p",
        "http://schemas.openxmlformats.org/presentationml/2006/main"
    )

    $shapeNodes =
        $SlideXml.SelectNodes(
            "//p:sp | //p:pic | //p:graphicFrame | //p:cxnSp | //p:grpSp",
            $ns
        )

    $shapes =
        @()

    $groups =
        @()

    $seenIds =
        @{}

    foreach ($shapeNode in $shapeNodes) {

        $cnv =
            Get-JungleCnvPrNode `
                -ShapeNode $shapeNode `
                -Ns $ns

        if ($null -eq $cnv) {

            continue
        }

        $shapeId =
            [string]$cnv.GetAttribute(
                "id"
            )

        if (
            [string]::IsNullOrWhiteSpace(
                $shapeId
            )
        ) {

            throw (
                "Shape without cNvPr id detected."
            )
        }

        if ($seenIds.ContainsKey(
            $shapeId
        )) {

            throw (
                "Duplicate shape id detected: " +
                $shapeId
            )
        }

        $seenIds[$shapeId] =
            $true

        $parentGroupId =
            $null

        $parent =
            $shapeNode.ParentNode

        while ($null -ne $parent) {

            if (
                [string]$parent.LocalName -eq
                "grpSp"
            ) {

                $parentCnv =
                    Get-JungleCnvPrNode `
                        -ShapeNode $parent `
                        -Ns $ns

                if ($null -ne $parentCnv) {

                    $parentGroupId =
                        [string]$parentCnv.GetAttribute(
                            "id"
                        )
                }

                break
            }

            if (
                [string]$parent.LocalName -eq
                "spTree"
            ) {

                break
            }

            $parent =
                $parent.ParentNode
        }

        $shapeInfo =
            [pscustomobject]@{

                id =
                    $shapeId

                kind =
                    [string]$shapeNode.LocalName

                name =
                    [string]$cnv.GetAttribute(
                        "name"
                    )

                descr =
                    [string]$cnv.GetAttribute(
                        "descr"
                    )

                parent_group_id =
                    $parentGroupId
            }

        $shapes +=
            $shapeInfo


        if (
            [string]$shapeNode.LocalName -eq
            "grpSp"
        ) {

            $memberNodes =
                $shapeNode.SelectNodes(
                    "./p:sp | ./p:pic | ./p:graphicFrame | ./p:cxnSp | ./p:grpSp",
                    $ns
                )

            $memberIds =
                @()

            foreach ($memberNode in $memberNodes) {

                $memberCnv =
                    Get-JungleCnvPrNode `
                        -ShapeNode $memberNode `
                        -Ns $ns

                if ($null -ne $memberCnv) {

                    $memberIds +=
                        [string]$memberCnv.GetAttribute(
                            "id"
                        )
                }
            }

            $groups +=
                [pscustomobject]@{

                    id =
                        $shapeId

                    members =
                        @(
                            $memberIds |
                            Sort-Object {
                                [int]$_
                            }
                        )
                }
        }
    }

    return [pscustomobject]@{

        shapes =
            @($shapes)

        groups =
            @($groups)
    }
}


function Get-JungleReflowPptSnapshot {

    param(
        [Parameter(Mandatory=$true)]
        [string]$Path
    )

    if (-not (
        Test-Path `
            -LiteralPath $Path `
            -PathType Leaf
    )) {

        throw (
            "PPTX missing: " +
            $Path
        )
    }

    $zip =
        [System.IO.Compression.ZipFile]::OpenRead(
            $Path
        )

    try {

        [xml]$presentationXml =
            Get-JungleZipEntryText `
                -Zip $zip `
                -EntryPath "ppt/presentation.xml"

        [xml]$presentationRels =
            Get-JungleZipEntryText `
                -Zip $zip `
                -EntryPath "ppt/_rels/presentation.xml.rels"


        $presentationNs =
            New-Object System.Xml.XmlNamespaceManager(
                $presentationXml.NameTable
            )

        $presentationNs.AddNamespace(
            "p",
            "http://schemas.openxmlformats.org/presentationml/2006/main"
        )

        $presentationNs.AddNamespace(
            "r",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        )


        $relMap =
            @{}

        $relNodes =
            $presentationRels.SelectNodes(
                "/*[local-name()='Relationships']/*[local-name()='Relationship']"
            )

        foreach ($relNode in $relNodes) {

            $relId =
                [string]$relNode.GetAttribute(
                    "Id"
                )

            $target =
                [string]$relNode.GetAttribute(
                    "Target"
                )

            if (
                -not [string]::IsNullOrWhiteSpace(
                    $relId
                )
            ) {

                $relMap[$relId] =
                    $target
            }
        }


        $slideIdNodes =
            $presentationXml.SelectNodes(
                "//p:sldIdLst/p:sldId",
                $presentationNs
            )

        $slides =
            @()

        $slideIds =
            @()


        foreach ($slideIdNode in $slideIdNodes) {

            $slideId =
                [string]$slideIdNode.GetAttribute(
                    "id"
                )

            $relId =
                [string]$slideIdNode.GetAttribute(
                    "id",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                )


            if (
                [string]::IsNullOrWhiteSpace(
                    $slideId
                )
            ) {

                throw (
                    "Presentation contains slide without stable slide id."
                )
            }


            if (
                [string]::IsNullOrWhiteSpace(
                    $relId
                )
            ) {

                throw (
                    "Presentation slide id missing relationship id."
                )
            }


            if (-not $relMap.ContainsKey(
                $relId
            )) {

                throw (
                    "Presentation relationship missing: " +
                    $relId
                )
            }


            $slidePart =
                Resolve-JungleZipRelativePath `
                    -BaseDirectory "ppt" `
                    -Target ([string]$relMap[$relId])


            [xml]$slideXml =
                Get-JungleZipEntryText `
                    -Zip $zip `
                    -EntryPath $slidePart


            $textInfo =
                Get-JungleReflowSemanticTextInfo `
                    -SlideXml $slideXml


            $shapeInfo =
                Get-JungleReflowShapeInfo `
                    -SlideXml $slideXml


            $slideIds +=
                $slideId


            $slides +=
                [pscustomobject]@{

                    slide_id =
                        $slideId

                    relationship_id =
                        $relId

                    part_path =
                        $slidePart

                    paragraph_count =
                        [int]$textInfo.paragraph_count

                    semantic_text_hash =
                        [string]$textInfo.semantic_text_hash

                    shapes =
                        @($shapeInfo.shapes)

                    groups =
                        @($shapeInfo.groups)
                }
        }


        return [pscustomobject]@{

            file_sha256 =
                Get-JungleSha256 `
                    -Path $Path

            slide_count =
                $slides.Count

            slide_ids =
                @($slideIds)

            slides =
                @($slides)
        }
    }
    finally {

        $zip.Dispose()
    }
}


function Compare-JungleReflowSnapshots {

    param(
        [Parameter(Mandatory=$true)]
        [object]$Source,

        [Parameter(Mandatory=$true)]
        [object]$Artifact
    )

    $violations =
        @()


    if (
        [int]$Source.slide_count -ne
        [int]$Artifact.slide_count
    ) {

        $violations +=
            "SLIDE_COUNT_CHANGED"
    }


    $sourceSlideIds =
        @(
            $Source.slide_ids |
            ForEach-Object {
                [string]$_
            }
        )

    $artifactSlideIds =
        @(
            $Artifact.slide_ids |
            ForEach-Object {
                [string]$_
            }
        )


    if (
        ($sourceSlideIds -join "|") -ne
        ($artifactSlideIds -join "|")
    ) {

        $violations +=
            "SLIDE_ORDER_OR_ID_CHANGED"
    }


    $artifactSlidesById =
        @{}

    foreach ($slide in @($Artifact.slides)) {

        $artifactSlidesById[
            [string]$slide.slide_id
        ] =
            $slide
    }


    foreach ($sourceSlide in @($Source.slides)) {

        $slideId =
            [string]$sourceSlide.slide_id


        if (-not $artifactSlidesById.ContainsKey(
            $slideId
        )) {

            $violations +=
                (
                    "SLIDE_MISSING:" +
                    $slideId
                )

            continue
        }


        $artifactSlide =
            $artifactSlidesById[$slideId]


        if (
            [int]$sourceSlide.paragraph_count -ne
            [int]$artifactSlide.paragraph_count
        ) {

            $violations +=
                (
                    "PARAGRAPH_COUNT_CHANGED:" +
                    $slideId
                )
        }


        if (
            [string]$sourceSlide.semantic_text_hash -ne
            [string]$artifactSlide.semantic_text_hash
        ) {

            $violations +=
                (
                    "TEXT_CHANGED:" +
                    $slideId
                )
        }


        $sourceShapes =
            @{}

        foreach ($shape in @($sourceSlide.shapes)) {

            $sourceShapes[
                [string]$shape.id
            ] =
                $shape
        }


        $artifactShapes =
            @{}

        foreach ($shape in @($artifactSlide.shapes)) {

            $artifactShapes[
                [string]$shape.id
            ] =
                $shape
        }


        foreach ($sourceShapeId in $sourceShapes.Keys) {

            if (-not $artifactShapes.ContainsKey(
                $sourceShapeId
            )) {

                $violations +=
                    (
                        "SOURCE_SHAPE_DELETED:" +
                        $slideId +
                        ":" +
                        $sourceShapeId
                    )

                continue
            }


            $sourceShape =
                $sourceShapes[$sourceShapeId]

            $artifactShape =
                $artifactShapes[$sourceShapeId]


            if (
                [string]$sourceShape.kind -ne
                [string]$artifactShape.kind
            ) {

                $violations +=
                    (
                        "SOURCE_SHAPE_KIND_CHANGED:" +
                        $slideId +
                        ":" +
                        $sourceShapeId
                    )
            }


            if (
                [string]$sourceShape.name -ne
                [string]$artifactShape.name
            ) {

                $violations +=
                    (
                        "SOURCE_SHAPE_NAME_CHANGED:" +
                        $slideId +
                        ":" +
                        $sourceShapeId
                    )
            }


            if (
                [string]$sourceShape.descr -ne
                [string]$artifactShape.descr
            ) {

                $violations +=
                    (
                        "SOURCE_SHAPE_DESCR_CHANGED:" +
                        $slideId +
                        ":" +
                        $sourceShapeId
                    )
            }


            if (
                [string]$sourceShape.parent_group_id -ne
                [string]$artifactShape.parent_group_id
            ) {

                $violations +=
                    (
                        "GROUP_MEMBERSHIP_CHANGED:" +
                        $slideId +
                        ":" +
                        $sourceShapeId
                    )
            }
        }


        foreach ($artifactShapeId in $artifactShapes.Keys) {

            if ($sourceShapes.ContainsKey(
                $artifactShapeId
            )) {

                continue
            }


            $newShape =
                $artifactShapes[$artifactShapeId]


            if (
                [string]$newShape.kind -eq
                "grpSp"
            ) {

                $violations +=
                    (
                        "NEW_GROUP_FORBIDDEN:" +
                        $slideId +
                        ":" +
                        $artifactShapeId
                    )

                continue
            }


            if (
                -not [string]::IsNullOrWhiteSpace(
                    [string]$newShape.parent_group_id
                )
            ) {

                $violations +=
                    (
                        "NEW_SHAPE_INSIDE_GROUP_FORBIDDEN:" +
                        $slideId +
                        ":" +
                        $artifactShapeId
                    )
            }


            if (-not (
                [string]$newShape.name
            ).StartsWith(
                "JUNGLE_REFLOW_DECOR_",
                [System.StringComparison]::Ordinal
            )) {

                $violations +=
                    (
                        "NEW_SHAPE_NAME_TAG_MISSING:" +
                        $slideId +
                        ":" +
                        $artifactShapeId
                    )
            }


            if (-not (
                [string]$newShape.descr
            ).Contains(
                "JUNGLE_REFLOW_V1=DECORATION"
            )) {

                $violations +=
                    (
                        "NEW_SHAPE_DESCR_TAG_MISSING:" +
                        $slideId +
                        ":" +
                        $artifactShapeId
                    )
            }
        }


        $sourceGroups =
            @{}

        foreach ($group in @($sourceSlide.groups)) {

            $sourceGroups[
                [string]$group.id
            ] =
                $group
        }


        $artifactGroups =
            @{}

        foreach ($group in @($artifactSlide.groups)) {

            $artifactGroups[
                [string]$group.id
            ] =
                $group
        }


        foreach ($sourceGroupId in $sourceGroups.Keys) {

            if (-not $artifactGroups.ContainsKey(
                $sourceGroupId
            )) {

                $violations +=
                    (
                        "SOURCE_GROUP_MISSING:" +
                        $slideId +
                        ":" +
                        $sourceGroupId
                    )

                continue
            }


            $sourceMembers =
                @(
                    $sourceGroups[$sourceGroupId].members |
                    ForEach-Object {
                        [string]$_
                    } |
                    Sort-Object {
                        [int]$_
                    }
                )


            $artifactMembers =
                @(
                    $artifactGroups[$sourceGroupId].members |
                    ForEach-Object {
                        [string]$_
                    } |
                    Sort-Object {
                        [int]$_
                    }
                )


            if (
                ($sourceMembers -join "|") -ne
                ($artifactMembers -join "|")
            ) {

                $violations +=
                    (
                        "GROUP_MEMBER_SET_CHANGED:" +
                        $slideId +
                        ":" +
                        $sourceGroupId
                    )
            }
        }


        foreach ($artifactGroupId in $artifactGroups.Keys) {

            if (-not $sourceGroups.ContainsKey(
                $artifactGroupId
            )) {

                $violations +=
                    (
                        "NEW_GROUP_FORBIDDEN:" +
                        $slideId +
                        ":" +
                        $artifactGroupId
                    )
            }
        }
    }


    return [pscustomobject]@{

        passed =
            ($violations.Count -eq 0)

        violation_count =
            $violations.Count

        violations =
            @($violations)
    }
}


function Assert-JungleReflowMutationBoundary {

    param(
        [Parameter(Mandatory=$true)]
        [string]$SourcePath,

        [Parameter(Mandatory=$true)]
        [string]$ArtifactPath,

        [Parameter(Mandatory=$true)]
        [string]$SourceHashBefore
    )


    $sourceFull =
        [System.IO.Path]::GetFullPath(
            $SourcePath
        )

    $artifactFull =
        [System.IO.Path]::GetFullPath(
            $ArtifactPath
        )


    if (
        [string]::Equals(
            $sourceFull,
            $artifactFull,
            [System.StringComparison]::OrdinalIgnoreCase
        )
    ) {

        throw (
            "REFLOW source overwrite is forbidden."
        )
    }


    $sourceHashAfter =
        Get-JungleSha256 `
            -Path $SourcePath


    if (
        $sourceHashAfter -ne
        $SourceHashBefore
    ) {

        throw (
            "REFLOW source file changed."
        )
    }


    $sourceSnapshot =
        Get-JungleReflowPptSnapshot `
            -Path $SourcePath


    $artifactSnapshot =
        Get-JungleReflowPptSnapshot `
            -Path $ArtifactPath


    $comparison =
        Compare-JungleReflowSnapshots `
            -Source $sourceSnapshot `
            -Artifact $artifactSnapshot


    if (-not [bool]$comparison.passed) {

        throw (
            "REFLOW Mutation Boundary V2 violation: " +
            (
                @($comparison.violations) -join
                " | "
            )
        )
    }


    return [pscustomobject]@{

        passed =
            $true

        version =
            "V2"

        source_hash_unchanged =
            $true

        slide_count =
            [int]$sourceSnapshot.slide_count

        violation_count =
            0

        violations =
            @()
    }
}