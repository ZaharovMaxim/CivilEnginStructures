Imports System.Collections.Generic
Imports System.IO
Imports System.Runtime.ExceptionServices
Imports NUnit.Framework

Namespace Tests
    <TestFixture>
    Public Class FuncXMLBugTests
        <Test>
        Public Sub FuncReadFieldValueDoesNotReadMatchingFieldFromAnotherTable()
            Dim xml As String =
                "<PropertyTables>" &
                "<PropertyTable TableName=""Wanted"">" &
                "<Field Name=""Other"" Value=""IGNORED"" />" &
                "</PropertyTable>" &
                "<PropertyTable TableName=""Wrong"">" &
                "<Field Name=""Target"" Value=""WRONG"" />" &
                "</PropertyTable>" &
                "</PropertyTables>"

            WithTemporaryXml(
                xml,
                Sub(path)
                    Dim values As String() = Nothing

                    Dim found As Boolean = FuncXML.FuncReadFieldValueByPSTablesToXML(
                        path,
                        "Wanted",
                        "Target",
                        values)

                    Assert.Multiple(
                        Sub()
                            Assert.That(found, [Is].False,
                                "The requested field is absent from the requested table.")
                            Assert.That(values, [Is].Null,
                                "A value from a different table must not be returned.")
                        End Sub)
                End Sub)
        End Sub

        <Test>
        Public Sub FuncReadFieldValueFindsValueWhenValueAttributePrecedesName()
            Dim xml As String =
                "<PropertyTables>" &
                "<PropertyTable TableName=""Wanted"">" &
                "<Field Value=""FIRST;SECOND"" Name=""Target"" />" &
                "</PropertyTable>" &
                "</PropertyTables>"

            WithTemporaryXml(
                xml,
                Sub(path)
                    Dim values As String() = Nothing

                    Dim found As Boolean = FuncXML.FuncReadFieldValueByPSTablesToXML(
                        path,
                        "Wanted",
                        "Target",
                        values)

                    Assert.Multiple(
                        Sub()
                            Assert.That(found, [Is].True)
                            Assert.That(values, [Is].EqualTo(New String() {"FIRST", "SECOND"}))
                        End Sub)
                End Sub)
        End Sub

        <Test>
        Public Sub ReadAlbumBeamsAddsEachBeamElementExactlyOnce()
            Dim xml As String =
                "<Beams>" &
                "<Beam model=""B1"" fullLenght=""12000"" />" &
                "</Beams>"

            WithTemporaryXml(
                xml,
                Sub(path)
                    Dim beamsByAlbum As New Dictionary(Of String, BeamI())()
                    Dim names As String() = Nothing

                    Dim read As Boolean = FuncXML.readAlbumBeamsFromXMLFiles(
                        path,
                        beamsByAlbum,
                        names,
                        "Album")
                    Dim beams As BeamI() = beamsByAlbum("Album")

                    Assert.Multiple(
                        Sub()
                            Assert.That(read, [Is].True)
                            Assert.That(beams.Length, [Is].EqualTo(1),
                                "One Beam XML element must create one BeamI object.")
                            Assert.That(names, [Is].EqualTo(New String() {"B1"}))
                            Assert.That(beams(0).model, [Is].EqualTo("B1"))
                            Assert.That(beams(0).lenght, [Is].EqualTo(12.0))
                        End Sub)
                End Sub)
        End Sub

        Private Shared Sub WithTemporaryXml(xml As String, assertion As Action(Of String))
            Dim temporaryFilePath As String = Path.Combine(
                TestContext.CurrentContext.WorkDirectory,
                Guid.NewGuid().ToString("N") & ".xml")
            Dim assertionFailure As Exception = Nothing
            Dim cleanupFailure As Exception = Nothing

            File.WriteAllText(temporaryFilePath, xml)

            Try
                assertion(temporaryFilePath)
            Catch ex As Exception
                assertionFailure = ex
            Finally
                Try
                    File.Delete(temporaryFilePath)
                Catch ex As Exception
                    cleanupFailure = ex
                End Try
            End Try

            If assertionFailure IsNot Nothing Then
                If cleanupFailure IsNot Nothing Then
                    TestContext.Progress.WriteLine(
                        "Temporary XML cleanup also failed: " & cleanupFailure.Message)
                End If
                ExceptionDispatchInfo.Capture(assertionFailure).Throw()
            End If

            If cleanupFailure IsNot Nothing Then
                ExceptionDispatchInfo.Capture(cleanupFailure).Throw()
            End If
        End Sub
    End Class
End Namespace
