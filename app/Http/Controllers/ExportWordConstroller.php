<?php

namespace App\Http\Controllers;

use Exception;
use Illuminate\Http\Request;

class ExportWordConstroller extends Controller
{
    public function execute(Request $request)
    {
        $wordTest = new \PhpOffice\PhpWord\PhpWord();

        $newSection = $wordTest->addSection();
        
        $desc1 = $request->konten ?? "";
        
        if(!$request->konten){
            $desc1 = "The Portfolio details is a very useful feature of the web page. You can establish your archived details and the works to the entire web community. It was outlined to bring in extra clients, get you selected based on this details.";                      
        }

        $newSection->addText($desc1, array('name' => 'Roboto', 'size' => 15));

        $objectWriter = \PhpOffice\PhpWord\IOFactory::createWriter($wordTest, 'Word2007');
        try {
            $objectWriter->save(storage_path('TestWordFile.docx'));
        } catch (Exception $e) {
        }
        return response()->download(storage_path('TestWordFile.docx'));
    }
}
