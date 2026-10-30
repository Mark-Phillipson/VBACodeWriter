Imports System.Collections.Generic
Imports System.Drawing
Imports Markdig
Imports Markdig.Syntax
Imports Markdig.Syntax.Inlines

Friend NotInheritable Class ChatMarkdown
    Public Shared Function GetCodeBlocks(markdown As String) As List(Of String)
        Dim snippets As New List(Of String)
        CollectCode(Markdig.Markdown.Parse(markdown), snippets)
        Return snippets
    End Function

    Private Shared Sub CollectCode(blocks As ContainerBlock, snippets As List(Of String))
        For Each block In blocks
            If TypeOf block Is CodeBlock Then
                Dim code = DirectCast(block, CodeBlock).Lines.ToString()
                If Not String.IsNullOrWhiteSpace(code) Then snippets.Add(code)
            ElseIf TypeOf block Is ContainerBlock Then
                CollectCode(DirectCast(block, ContainerBlock), snippets)
            End If
        Next
    End Sub

    Public Shared Sub AppendResponse(view As RichTextBox, markdown As String)
        Dim responseStart = view.TextLength
        RenderBlocks(view, Markdig.Markdown.Parse(markdown), 0)
        view.Select(responseStart, 0)
        view.ScrollToCaret()
    End Sub

    Public Shared Sub AppendText(view As RichTextBox, text As String, Optional style As FontStyle = FontStyle.Regular,
                                 Optional code As Boolean = False, Optional size As Single = 10.0F)
        view.SelectionStart = view.TextLength
        view.SelectionLength = 0
        Using textFont As New Font(If(code, "Consolas", "Segoe UI"), size, style)
            view.SelectionFont = textFont
            view.SelectionColor = Color.FromArgb(32, 35, 39)
            view.SelectionBackColor = If(code, Color.FromArgb(240, 243, 246), Color.White)
            view.AppendText(text)
        End Using
    End Sub

    Private Shared Sub RenderBlocks(view As RichTextBox, blocks As ContainerBlock, depth As Integer)
        For Each block In blocks
            If TypeOf block Is CodeBlock Then
                AppendText(view, DirectCast(block, CodeBlock).Lines.ToString() & vbCrLf & vbCrLf, code:=True)
            ElseIf TypeOf block Is HeadingBlock Then
                Dim heading = DirectCast(block, HeadingBlock)
                RenderInlines(view, heading.Inline, FontStyle.Bold, If(heading.Level <= 2, 13.0F, 11.0F))
                AppendText(view, vbCrLf & vbCrLf)
            ElseIf TypeOf block Is ListBlock Then
                Dim list = DirectCast(block, ListBlock)
                Dim itemNumber As Integer = 1
                If list.IsOrdered Then Integer.TryParse(list.OrderedStart, itemNumber)
                For Each item As ListItemBlock In list
                    AppendText(view, New String(" "c, depth * 2) & If(list.IsOrdered, itemNumber.ToString() & ". ", "- "))
                    RenderBlocks(view, item, depth + 1)
                    itemNumber += 1
                Next
            ElseIf TypeOf block Is QuoteBlock Then
                AppendText(view, "> ", FontStyle.Italic)
                RenderBlocks(view, DirectCast(block, ContainerBlock), depth + 1)
            ElseIf TypeOf block Is ContainerBlock Then
                RenderBlocks(view, DirectCast(block, ContainerBlock), depth)
            ElseIf TypeOf block Is ThematicBreakBlock Then
                AppendText(view, "--------------------" & vbCrLf)
            ElseIf TypeOf block Is LeafBlock Then
                Dim leaf = DirectCast(block, LeafBlock)
                If leaf.Inline IsNot Nothing Then
                    RenderInlines(view, leaf.Inline, FontStyle.Regular, 10.0F)
                Else
                    AppendText(view, leaf.Lines.ToString())
                End If
                AppendText(view, vbCrLf & vbCrLf)
            End If
        Next
    End Sub

    Private Shared Sub RenderInlines(view As RichTextBox, container As ContainerInline, style As FontStyle, size As Single)
        If container Is Nothing Then Return
        Dim inline As Inline = container.FirstChild
        Do While inline IsNot Nothing
            If TypeOf inline Is LiteralInline Then
                AppendText(view, DirectCast(inline, LiteralInline).Content.ToString(), style, size:=size)
            ElseIf TypeOf inline Is CodeInline Then
                AppendText(view, DirectCast(inline, CodeInline).Content, style, code:=True, size:=size)
            ElseIf TypeOf inline Is LineBreakInline Then
                AppendText(view, vbCrLf)
            ElseIf TypeOf inline Is EmphasisInline Then
                Dim emphasis = DirectCast(inline, EmphasisInline)
                RenderInlines(view, emphasis, style Or If(emphasis.DelimiterCount >= 2, FontStyle.Bold, FontStyle.Italic), size)
            ElseIf TypeOf inline Is LinkInline Then
                Dim link = DirectCast(inline, LinkInline)
                RenderInlines(view, link, style Or FontStyle.Underline, size)
                If Not String.IsNullOrEmpty(link.Url) Then AppendText(view, " (" & link.Url & ")", size:=size)
            ElseIf TypeOf inline Is AutolinkInline Then
                AppendText(view, DirectCast(inline, AutolinkInline).Url, size:=size)
            ElseIf TypeOf inline Is HtmlInline Then
                AppendText(view, DirectCast(inline, HtmlInline).Tag, size:=size)
            ElseIf TypeOf inline Is ContainerInline Then
                RenderInlines(view, DirectCast(inline, ContainerInline), style, size)
            End If
            inline = inline.NextSibling
        Loop
    End Sub
End Class