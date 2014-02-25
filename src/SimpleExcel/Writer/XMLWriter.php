<?php

namespace SimpleExcel\Writer;

/**
 * SimpleExcel class for writing Microsoft Excel 2003 XML Spreadsheet
 *
 * @author  Faisalman
 * @package SimpleExcel
 */
class XMLWriter extends BaseWriter implements IWriter
{
    /**
     * Defines content-type for HTTP header
     *
     * @access  protected
     * @var     string
     */
    protected $content_type = 'application/xml';

    /**
     * Defines file extension to be used when saving file
     *
     * @access  protected
     * @var     string
     */
    protected $file_extension = 'xml';

    /**
     * Array containing document properties
     *
     * @access  private
     * @var     array
     */
    private $doc_prop;
    
    /**
     * Array containing document styles
     *
     * @access  private
     * @var     array
     */
    private $doc_styles;

    /**
     * @return  void
     */
    public function __construct(){
        $this->doc_prop = array(
            'Author' => 'SimpleExcel',
            'Company' => 'SimpleExcel',
            'Created' => gmdate("Y-m-d\TH:i:s\Z"),
            'Keywords' => 'SimpleExcel',
            'LastAuthor' => 'SimpleExcel',
            'Version' => '12.00'
        );
        $this->doc_prop = array();
    }

    /**
     * Adding row data to XML
     *
     * @param   array   $values An array contains ordered value for every cell
     * @return  void
     */
    public function addRow($values){
        $row = &$this->tabl_data;
        $row .= '
    <Row ss:AutoFitHeight="0">';

        foreach($values as $val){

            $value = '';
            $datatype = 'String';
            $cell_attributes = '';

            // check if given variable contains array
            if(is_array($val) && isset($val['value'])){
                $value = $val['value'];
                $datatype = isset($val['datatype']) ? $val['datatype'] : is_string($val['value']) ? 'String' : (is_numeric($val['value']) ? 'Number' : 'String');
                $cell_attributes = isset($val['cell_attributes']) ? $val['cell_attributes'] : '';
            } elseif(is_array($val)){
                $value = $val[0];
                $datatype = $val[1];
            } else {
                $value = $val;
                $datatype = is_string($val) ? 'String' : (is_numeric($val) ? 'Number' : 'String');
            }

            if (!empty($cell_attributes)) {
                $cell_attributes = ' ' . $cell_attributes;
            }

            // escape value from HTML tags
            $value = filter_var($value, FILTER_SANITIZE_SPECIAL_CHARS);

            $row .= '
    <Cell'.$cell_attributes.'><Data ss:Type="'.$datatype.'">'.$value.'</Data></Cell>';
        }

        $row .= '
    </Row>';
    }

    /**
     * Get document content as string
     *
     * @return  string  Content of document
     */
    public function saveString(){
        $content = '<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">';

        foreach($this->doc_prop as $propname => $propval){
            $content .= '
  <'.$propname.'>'.$propval.'</'.$propname.'>';
        }

  $content .= '
 </DocumentProperties>';
  
  if (count($this->doc_styles)){
    $content .= '
 <Styles>';
    foreach ($this->doc_styles as $style_content){
      $content .= '
  ' . $style_content;
    }
    $content .= '
 </Styles>';
  }
        
 $content .= '
 <Worksheet ss:Name="Sheet1">
  <Table>'.$this->tabl_data.'
  </Table>
 </Worksheet>
</Workbook>';
        return $content;
    }

    /**
    * Set XML data
    *
    * @param    array   $values An array contains ordered value of arrays for all fields
    * @return   void
    */
    public function setData($values){
        if(!is_array($values)){
            $values = array($values);
        }
        $this->tabl_data = ""; // reset the xml data.

        // append values as rows
        foreach ($values as $value) {
            $this->addRow($value);
        }
    }

    /**
    * Set a document property of the XML
    *
    * @param    string  $prop   Document property to be set
    * @param    string  $val    Value of the document property
    * @return   void
    */
    public function setDocProp($prop, $val){
        $this->doc_prop[$prop] = $val;
    }
    
    /**
    * Set a document style
    *
    * @param    string  $style_id           The style id
    * @param    string  $style_elements     Style elements
    * @return   void
    */
    public function setDocStyle($style_id, $style_elements){
        $style_elements_content = '';
        if (is_array($style_elements) && count($style_elements)){
          foreach ($style_elements as $style_element_name => $style_element_properties){
            $properties = array();
            if (is_array($style_element_properties) && count($style_element_properties)){
              foreach ($style_element_properties as $property_name => $property_value){
                $properties[] = 'ss:' . $property_name .'="' . $property_value . '"';
              }
            }
            $style_elements_content .= '
     <' . $style_element_name . ' ' . implode(' ', $properties) . '/>';
          }
        }
        if (!empty($style_elements_content)){
          $this->doc_styles[] = '<Style ss:ID="' . $style_id . '">' . $style_elements_content . '
  </Style>';
        }
    }
}
?>
