Option Explicit

    '�h���b�O�A���h�h���b�v�Ŏ擾�����t�@�C���p�X��ϐ��ɓ����
    Dim GetPathArray
    Set GetPathArray = WScript.Arguments
    
    '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    '�C�e���[�^
    Dim pt

    '�t�@�C���̐��Ԃ񃋁[�v����
    For Each pt in GetPathArray

        '�擾�����t�@�C����
        Dim FileName
        FileName = objFSO.GetFileName(pt)

        '�t�@�C���̃t���p�X��InputBox�ŕ\��������
        pt = InputBox("�h���b�O�A���h�h���b�v�����t�@�C���� " & FileName,"�t�@�C���̃t���p�X��\��", pt)

    Next

    '�I�u�W�F�N�g�ϐ����N���A
    Set objFSO = Nothing
