<?php
/**
 * Created by PhpStorm.
 * User: assadmahmood
 * Date: 30/06/15
 * Time: 10:45 AM
 */

namespace Aspose\Cells\QuickStart;


class HelloWorld {

    public static function run()
    {
        $ptr= new \COM('Aspose.Cells.Interop.InteropHelper');



        $doc = new \COM("Aspose.Words.Document");

        $builder = new \COM("Aspose.Words.DocumentBuilder");

        $builder->Document = $doc;

        $builder->Write("Hello world!");

        $doc->Save("./data/HelloWorld Out.docx");


        /*$ptr= new \COM("Aspose.Words.ComHelper")or die('Unable to load helper');

        $doc = $ptr->New("Document",array());

        $builder = $ptr->New("Aspose.Words.DocumentBuilder",array($doc));

        $builder->Writeln("Hello World!");

        $doc->Save("./data/HelloWorld Out.docx");*/


    }

} 