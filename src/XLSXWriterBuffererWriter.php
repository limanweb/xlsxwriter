<?php 

namespace Limanweb\XLSXWriter;

class XLSXWriterBuffererWriter
{
    protected $fd = null;
    protected $buffer = '';
    protected $checkUTF8 = false;
    
    /**
     * 
     * @param string $fileName
     * @param string $fdFopenFlags
     * @param boolean $checkUTF8
     */
    public function __construct($fileName, $fdFopenFlags = 'w', $checkUTF8 = false)
    {
        $this->checkUTF8 = $checkUTF8;
        $this->fd = fopen($fileName, $fdFopenFlags);
        if ($this->fd===false) {
            XLSXWriter::log("Unable to open $fileName for writing.");
        }
    }
    
    /**
     * 
     * @param string $string
     */
    public function write($string)
    {
        $this->buffer.=$string;
        if (isset($this->buffer[8191])) {
            $this->purge();
        }
    }
    
    /**
     * 
     */
    protected function purge()
    {
        if ($this->fd) {
            if ($this->checkUTF8 && !self::isValidUTF8($this->buffer)) {
                XLSXWriter::log("Error, invalid UTF8 encoding detected.");
                $this->checkUTF8 = false;
            }
            fwrite($this->fd, $this->buffer);
            $this->buffer = '';
        }
    }
    
    /**
     * 
     */
    public function close()
    {
        $this->purge();
        if ($this->fd) {
            fclose($this->fd);
            $this->fd = null;
        }
    }
    
    /**
     * 
     */
    public function __destruct()
    {
        $this->close();
    }
    
    /**
     * 
     */
    public function ftell()
    {
        if ($this->fd) {
            $this->purge();
            return ftell($this->fd);
        }
        return -1;
    }
    
    /**
     * 
     * @param number $pos
     * @return number
     */
    public function fseek($pos)
    {
        if ($this->fd) {
            $this->purge();
            return fseek($this->fd, $pos);
        }
        return -1;
    }
    
    protected static function isValidUTF8($string)
    {
        if (function_exists('mb_check_encoding'))
        {
            return mb_check_encoding($string, 'UTF-8') ? true : false;
        }
        return preg_match("//u", $string) ? true : false;
    }
}
