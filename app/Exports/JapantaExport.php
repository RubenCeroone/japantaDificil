<?php

namespace App\Exports;

use App\Models\Japanta;
use Illuminate\Support\Collection;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithStyles;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use Carbon\Carbon;

class JapantaExport implements FromCollection, WithStyles
{
    protected $japanta;

    public function __construct()
    {
        $this->japanta = Japanta::all();
    }

    public function collection(): Collection
    {
        return $this->japanta;
    }

    public function clearRow1(Worksheet $sheet)
    {
        // Borra los datos de la línea 1 desde la columna B hasta la columna H
        for ($column = 'B'; $column <= 'H'; $column++) {
            $sheet->setCellValue($column . '1', '');
        }
    }

    public function clearRow2(Worksheet $sheet)
    {
        // Borra los datos desde la celda C2 hasta la celda H6
        for ($row = 2; $row <= 6; $row++) {
            for ($column = 'C'; $column <= 'H'; $column++) {
                $sheet->setCellValue($column . $row, '');
            }
        }
    }

    public function clearRow3(Worksheet $sheet)
    {
        // Borra los datos desde la celda A5 hasta la celda B6
        for ($row = 5; $row <= 6; $row++) {
            for ($column = 'A'; $column <= 'B'; $column++) {
                $sheet->setCellValue($column . $row, '');
            }
        }
    }

    public function clearRow4(Worksheet $sheet)
    {
        // Borra los datos desde la celda H7 hasta la celda H100
        for ($row = 7; $row <= 100; $row++) {
            $sheet->setCellValue('H' . $row, '');
        }
    }

    public function styles(Worksheet $sheet)
    {
        // Llamar a la función clearRow1
        $this->clearRow1($sheet);
        // Llamar a la función clearRow2
        $this->clearRow2($sheet);
        // Llamar a la función clearRow3
        $this->clearRow3($sheet);
        // Llamar a la función clearRow4
        $this->clearRow4($sheet);

        // Establecer el ancho de las columnas
        $sheet->getColumnDimension('A')->setWidth(12.86);
        $sheet->getColumnDimension('B')->setWidth(59.00);
        $sheet->getColumnDimension('C')->setWidth(22);
        $sheet->getColumnDimension('D')->setWidth(13.43);
        $sheet->getColumnDimension('E')->setWidth(13.43);
        $sheet->getColumnDimension('F')->setWidth(13.43);
        $sheet->getColumnDimension('G')->setWidth(13.43);

        // Establecer el alto de la fila 2 en 24
        $sheet->getRowDimension(2)->setRowHeight(24);

        // Aplicar estilos a la celda A1
        $sheet->getStyle('A1')->getFont()->setBold(true)->setSize(12);
        $sheet->getStyle('A1')->applyFromArray([
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_LEFT,
                'vertical' => Alignment::VERTICAL_CENTER,
            ],
        ]);

        // Escribe el texto en la celda A1
        $sheet->setCellValue('A1', 'Japanta SL');

        // Combina las celdas A2 y B2
        $sheet->mergeCells('A2:B2');

        // Aplica estilos al texto en la celda A2
        $sheet->getStyle('A2')->getFont()->setBold(true)->setSize(18);
        $sheet->getStyle('A2')->applyFromArray([
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_LEFT,
                'vertical' => Alignment::VERTICAL_CENTER,
            ],
        ]);

        // Escribe el texto en la celda A2
        $sheet->setCellValue('A2', 'Japanta SL - Libro Mayor 01/09/2023-30/09/2023');

        // Combina las celdas A3 y B3
        $sheet->mergeCells('A3:B3');

        // Aplica estilos al texto en la celda A3
        $sheet->getStyle('A3')->getFont()->setSize(11);
        $sheet->getStyle('A3')->applyFromArray([
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_LEFT,
                'vertical' => Alignment::VERTICAL_CENTER,
            ],
        ]);

        // Escribe el texto en la celda A3
        $sheet->setCellValue('A3', '01/09/2023 - 30/09/2023');

        // Combina las celdas A4 y B4
        $sheet->mergeCells('A4:B4');

        // Aplica estilos al texto en la celda A4
        $sheet->getStyle('A4')->getFont()->setSize(11);
        $sheet->getStyle('A4')->applyFromArray([
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_LEFT,
                'vertical' => Alignment::VERTICAL_CENTER,
            ],
        ]);

        // Escribe el texto en la celda A4
        $sheet->setCellValue('A4', 'añadir desde/hasta de la sección de cuentas contables');

        // Establece el color de fondo de las celdas A7:G7 a negro
        $sheet->getStyle('A7:G7')->applyFromArray([
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => [
                    'argb' => '000000', // Código ARGB para negro
                ],
            ],
        ]);

        // Aplica estilos al texto de las celdas A7:G7
        $sheet->getStyle('A7:G7')->getFont()->setSize(11);
        $sheet->getStyle('A7:G7')->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['argb' => 'FFFFFFFF'], // Color blanco
            ],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_LEFT,
                'vertical' => Alignment::VERTICAL_CENTER,
            ],
        ]);

        // Establece el texto en las celdas A7:G7
        $sheet->setCellValue('A7', 'Cuenta');
        $sheet->setCellValue('B7', 'Nombre');
        $sheet->setCellValue('C7', 'Grupo');
        $sheet->setCellValue('D7', 'Saldo Inicial');
        $sheet->setCellValue('E7', 'Debe');
        $sheet->setCellValue('F7', 'Haber');
        $sheet->setCellValue('G7', 'Saldo Final');

        // Insertar datos de Japanta y Japanta1
        $startRow = 8; // Fila donde empiezan los datos
        $endRow = 100; // Última fila donde se insertarán datos

        $row = $startRow;
        $totalDebe = 0;
        $totalHaber = 0;
        foreach ($this->collection() as $japanta) {
            if ($row > $endRow) {
                break; // Salir del bucle si alcanzamos la última fila
            }
            $sheet->setCellValue('A' . $row, $japanta->cuenta);
            $sheet->getStyle('A' . $row)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT);
            $sheet->setCellValue('B' . $row, $japanta->nombre);
            $sheet->getStyle('B' . $row)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
            $sheet->setCellValue('C' . $row, $japanta->grupo);
            $sheet->getStyle('C' . $row)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
            $sheet->setCellValue('D' . $row, $japanta->saldoinicial);

            // Formatear números en las columnas debe, haber y saldofinal
            $sheet->setCellValue('E' . $row, $japanta->debe != 0 ? number_format($japanta->debe, 2) : '');
            $sheet->setCellValue('F' . $row, $japanta->haber != 0 ? number_format($japanta->haber, 2) : '');

            // Sumar debe y haber y colocar el resultado en saldofinal (columna G)
            $saldofinal = $japanta->debe - $japanta->haber;
            $sheet->setCellValue('G' . $row, $saldofinal != 0 ? number_format($saldofinal, 2) : '');

            // Aplicar estilos y alinear a la derecha el contenido de las celdas de las columnas debe, haber y saldo
            $sheet->getStyle('E' . $row)->getNumberFormat()->setFormatCode('#,##0.00');
            $sheet->getStyle('E' . $row)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT);

            $sheet->getStyle('F' . $row)->getNumberFormat()->setFormatCode('#,##0.00');
            $sheet->getStyle('F' . $row)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT);

            $sheet->getStyle('G' . $row)->getNumberFormat()->setFormatCode('#,##0.00');
            $sheet->getStyle('G' . $row)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT);

            $row++;
        }
    }
}