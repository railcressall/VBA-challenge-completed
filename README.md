# VBA-challenge-completed
completed VBA Module 2 Challenge


** folowing code under 'math in my script was written with the assistance of a tutor. Hassan

                If ws.Cells(start, 3) = 0 Then
                    For find_value = start To i
                        If ws.Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                            Exit For
                        End If
                    Next find_value
                End If
