# r2soft/excel

Geração de planilhas XLSX em PHP puro, sem dependências. É uma versão adaptada do [mk-j/PHP_XLSXWriter](https://github.com/mk-j/PHP_XLSXWriter), com formatação de números e datas por coluna.

```bash
composer require r2soft/excel
```

```php
use R2Soft\Excel\XLSXWriter;

$writer = new XLSXWriter();
$writer->writeSheetHeader('Planilha1', ['Nome' => 'string', 'Valor' => '#,##0.00']);
$writer->writeSheetRow('Planilha1', ['Fulano', 1234.5]);
$writer->writeToFile('relatorio.xlsx');   // ou writeToString() / writeToStdOut()
```

Licença MIT (veja `LICENSE`), mantendo o copyright do projeto original.
