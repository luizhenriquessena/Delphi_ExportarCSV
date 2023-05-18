unit ExportCsv;

interface

uses
  ShellAPI, DBGrids, Dialogs, Classes, DB, SysUtils, Windows;

  procedure ExportarCSVDataSet(Sender: TObject; Grid: TDBGrid);
  procedure ExportarCSVDBGrid(Sender: TObject; Grid: TDBGrid);

implementation

procedure ExportarCSVDataSet(Sender: TObject;
  Grid: TDBGrid);
var
  l, c: Integer;
  linha: String;
  SaveDialog: TSaveDialog;
  StringList: TStringList;
  DataSet: TDataSet;
  Field: TField;
  nomeArquivo: String;
begin
  SaveDialog := TSaveDialog.Create(nil);
  StringList := TStringList.Create;
  try
    SaveDialog.Filter := 'Arquivos CSV (*.csv)|*.csv';
    if SaveDialog.Execute then
    begin
      DataSet := Grid.DataSource.DataSet;

      // Criar o cabeçalho no StringList
      for c:= 0 to Grid.Columns.Count -1 do
      begin
        Field := DataSet.FieldByName(Grid.Columns[c].FieldName);
        linha := linha + Field.DisplayLabel + ';'; // ; é o separador de colunas
      end;
      StringList.Add(linha);

      // Adicionar as linhas no StringList
      DataSet.First;
      while not DataSet.Eof do
      begin
        linha := '';
        for c := 0 to Grid.Columns.Count -1 do
        begin
          Field := DataSet.FieldByName(Grid.Columns[c].FieldName);
          linha := linha + Field.AsString + ';';
        end;
        StringList.Add(linha);
        DataSet.Next;
      end;

      // Gerar arquivo CSV
      nomeArquivo := ChangeFileExt(SaveDialog.FileName, '.csv');
      StringList.SaveToFile(nomeArquivo);

      ShowMessage('Lista exportada com sucesso!');

      // Abrir o arquivo CSV com o aplicativo padrão do sistema
      ShellExecute(0, 'open', PChar(SaveDialog.FileName), nil, nil, SW_SHOWNORMAL);
    end;
  finally
    SaveDialog.Free;
    StringList.Free;
  end;
end;

procedure ExportarCSVDBGrid(Sender: TObject;
  Grid: TDBGrid);
var
  l, c: Integer;
  linha: string;
  SaveDialog: TSaveDialog;
  StringList: TStringList;
  DataSet: TDataSet;
  nomeArquivo: String;
begin
  SaveDialog := TSaveDialog.Create(nil);
  StringList := TStringList.Create;
  try
    SaveDialog.Filter := 'Arquivos CSV (*.csv)|*.csv';
    if SaveDialog.Execute then
    begin
      DataSet := Grid.DataSource.DataSet;

      // Percorrer as colunas do DBGrid
      for c := 0 to Grid.Columns.Count - 1 do
      begin
        // Verificar se a coluna está visível
        if Grid.Columns[c].Visible then
        begin
          linha := linha + Grid.Columns[c].Title.Caption + ';'; // Use ';' como separador CSV
        end;
      end;
      StringList.Add(linha);

      // Percorrer as linhas do DataSet
      DataSet.First;
      while not DataSet.Eof do
      begin
        linha := '';
        for c := 0 to Grid.Columns.Count - 1 do
        begin
          // Verificar se a coluna está visível
          if Grid.Columns[c].Visible then
          begin
            linha := linha + DataSet.FieldByName(Grid.Columns[c].FieldName).AsString + ';';
          end;
        end;
        StringList.Add(linha);
        DataSet.Next;
      end;

      // Gerar arquivo CSV
      nomeArquivo := ChangeFileExt(SaveDialog.FileName, '.csv');
      StringList.SaveToFile(nomeArquivo);

      ShowMessage('Exportação concluída com sucesso!');

      // Abrir o arquivo CSV com o aplicativo padrão do sistema
      ShellExecute(0, 'open', PChar(SaveDialog.FileName), nil, nil, SW_SHOWNORMAL);
    end;
  finally
    SaveDialog.Free;
    StringList.Free;
  end;
end;

end.
