<?php if (!defined('BASEPATH')) exit('No direct script access allowed');

// Export to Excel
// 
// This method generates an excel file and then returns the generated content

if (!function_exists('arrayToExcel'))
{
    function arrayToExcel($query, $fields, $filename = "Excel"){

        if (count($query) == 0) {
            return "The query is empty. It doesn't have any data.";
        } else {
            $headers = "";
            foreach ($fields as $field) {
                $headers .= $field . "\t";
            }

            $data = "";
            foreach ($query as $row) {
                $line = "";
                foreach ($row as $value) {
                    if ((!isset($value)) || ($value == "")) {
                        $value = "\t";
                    } else {
                        $value = str_replace('"', '""', $value);
                        $value = '"' . utf8_decode($value) . '"' . "\t";
                    }
                    $line .= $value;
                }
                $data .= trim($line) . "\n";
            }
            $data = str_replace("\r", "", $data);

            $content = $headers . "\n" . $data;
            $filename = date('YmdHis') . "_export_{$filename}.xls";

            header("Content-type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            header("Content-Disposition: attachment; filename={$filename}");
            header("Content-Length: " . strlen($content));
            header("Pragma: no-cache");

            return $content;
        }
    }
}

/* End of file export_excel_helper.php */
/* Location: ./application/helpers/export_excel_helper.php */