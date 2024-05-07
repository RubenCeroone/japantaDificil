<?php

namespace App\Http\Controllers;

use App\Http\Controllers\Controller;
use App\Models\Japanta;
use Maatwebsite\Excel\Facades\Excel;
use App\Exports\JapantaExport;

class BotonExcelController extends Controller
{
    public function invoke()
    {
        return view('botonexcel');
    }

    public function exportarExcel()
    {
        $datos = Japanta::all();
        return Excel::download(new JapantaExport($datos), 'Japanta.xlsx');
    }
}