<?php
/**
 * WordTemplateEngine - creation of WORD documents from .docx templates,
 * conversion of the documents through LibreOffice to other formats: PDF, HTML, XHTML, HTML adapted to email.
 *
 * PHP Version 7.1
 *
 * @see        https://github.com/philip-sorokin/word-template-engine The WordTemplateEngine GitHub project
 * @see        https://addondev.com/opensource/word-template-engine The project manual
 *
 * @version    1.0
 * @author     Philip Sorokin <philip.sorokin@gmail.com>
 * @copyright  2021 - Philip Sorokin
 * @license    http://www.gnu.org/licenses/gpl-3.0.html GNU General Public License
 * @note       This program is distributed in the hope that it will be useful - WITHOUT
 * ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
 */

/**
 * WordTemplateEngineExamples is a sample class for demonstrating WordTemplateEngine methods.
 *
 * @author   Philip Sorokin <philip.sorokin@gmail.com>
 * @see      https://addondev.com/manuals/word-template-engine The project manual
 */
class WordTemplateEngineExamples
{
	/**
	 * Template getter example.
	 *
	 * @static
	 *
	 * @return   WordTemplateEngine
	 *
	 */
	public static function getTemplate(): WordTemplateEngine
	{
		$templatePath = __DIR__ . '/template.docx';

		$tmpDir_optional = $_SERVER['DOCUMENT_ROOT'];
		$errorHandler_optional  = [__CLASS__, 'self::errorHandler'];

		return new WordTemplateEngine($templatePath, $tmpDir_optional, $errorHandler_optional);
	}


	/**
	 * Example of variable substitution.
	 *
	 * @param    WordTemplateEngine   $template
	 * @static
	 *
	 * @return   WordTemplateEngine
	 *
	 */
	public static function replaceVariables(WordTemplateEngine $template): WordTemplateEngine
	{
		// Clone the table row with the variable ${qty} five times
		$template->cloneRow('qty', 5);

		// The row variables
		$products = [
			1 => [
				'qty' => 3,
				'item_sku' => '1-ATF-5',
				'description' => 'Full Arch Bed Grey',
				'unit_price' => '500 €',
				'discount' => '15 €',
				'total_item' => '1485 €',
			],
			2 => [
				'qty' => 1,
				'item_sku' => '4-YMJ-7',
				'description' => 'Night Stand White',
				'unit_price' => '200 €',
				'discount' => '10 €',
				'total_item' => '190 €',
			],
			3 => [
				'qty' => 2,
				'item_sku' => '15-IMA-6',
				'description' => 'Mirror Driftwood',
				'unit_price' => '400 €',
				'discount' => '',
				'total_item' => '800 €',
			],
			4 => [
				'qty' => 4,
				'item_sku' => '5-MJ8-16',
				'description' => 'Dresser Grey',
				'unit_price' => '600 €',
				'discount' => '',
				'total_item' => '2400 €',
			],
			5 => [
				'qty' => 2,
				'item_sku' => '9-ODA-8',
				'description' => 'Twin Bunk Driftwood',
				'unit_price' => '800 €',
				'discount' => '20 €',
				'total_item' => '1580 €',
			],
		];

		foreach($products as $key => $product)
		{
			foreach($product as $prop => $value)
			{
				$name = $prop . '#' . $key;
				$template->setValue($name, $value);
			}
		}

		// Other variables
		$orderInfo = [
			'subtotal' => '6455 €',
			'sales_tax' => '700 €',
			'total_discount' => '45 €',
			'total_sum' => '7155 €',
			'salesperson' => 'Jane Doe',
			'job' => 'Sales assistant',
			'shipping_method' => 'FedEx',
			'shipping_terms' => 'Delivered At Place',
			'delivery_date' => '2030/12/31',
			'payment_terms' => 'Cash on delivery',
			'due_date' => '2030/12/30',
			'to_name' => 'Elisha Barton',
			'to_company_name' => 'Ondricka LLC',
			'to_street_address' => '375 Nikko Fall Apt. 358',
			'to_phone' => '+19(440)554-5942',
			'ship_to_name' => 'Ila Gutmann',
			'ship_to_company_name' => 'Bartoletti Group',
			'ship_to_street_address' => '1707 Berge Viaduct Suite 724',
			'ship_to_phone' => '+13813440037',
			'invoice_date' => '2030/12/01',
			'invoice_id' => 'A19583',
			'customer_id' => 'U385',
			'company_name' => 'Witting Group',
			'company_director' => 'John Doe',
			'company_address' => '1250 Lula River Suite 965',
			'company_phone' => '+1-653-572-8295',
			'company_fax' => '+1-653-572-8787',
			'company_email' => 'info@skiles.info',
		];

		foreach($orderInfo as $name => $value)
		{
			$template->setValue($name, $value);
		}

		return $template;
	}


	/**
	 * Example of alternative syntax variable substitution.
	 *
	 * @param    WordTemplateEngine   $template
	 * @static
	 *
	 * @return   WordTemplateEngine
	 *
	 */
	public static function replaceAlternativeVariables(WordTemplateEngine $template): WordTemplateEngine
	{
		// Switch to the alternative variable syntax like ~(var_name) if you like it more, it also allows to replace the variables in links.
		$template->alternativeSyntax(true);

		$orderInfo = [
			'company_name' => 'AddonDev',
			'company_phone' => '+79264104108',
			'company_email' => 'philip.sorokin@gmail.com',
			'company_website' => 'https://addondev.com',
			'github_url' => 'https://github.com/philip-sorokin/word-template-engine',
			'donate_url' => 'https://addondev.com/donate',
		]; 

		foreach($orderInfo as $name => $value)
		{
			$template->setValue($name, $value);
		}

		// Switch back to default syntax like ${var_name} if you need.
		$template->alternativeSyntax(false);

		return $template;
	}


	/**
	 * Example of defining the document metadata.
	 *
	 * @param    WordTemplateEngine   $template
	 * @static
	 *
	 * @return   WordTemplateEngine
	 *
	 */
	public static function setDocumentInfo(WordTemplateEngine $template): WordTemplateEngine
	{
		// Drop all metadata 
		// Warning! After calling this method you have to redefine the document creation time, the title and the document creator!
		$template->dropMetaData();

		// Set new metadata
		$template->setTime();
		$template->setCompany('Witting Group');
		$template->setManager('John Doe');
		$template->setTitle('Invoice A19583');
		$template->setAuthor('Jane Doe');
		$template->setSubject('Invoice for Elisha Barton');
		$template->setKeywords('Documents, payment, order');

		// Delete this metadata (actually, it's already removed, we use this as an example).
		$template->setDescription('');
		$template->setCategory('');
		$template->setStatus('');

		return $template;
	}


	/**
	 * Example of outputting a document in DOCX format.
	 *
	 * @static
	 *
	 * @return  void
	 *
	 */
	public static function outputDocument(): void
	{
		// Get the template.
		$template = self::getTemplate();

		// Set document metadata.
		$template = self::setDocumentInfo($template);

		// Replace the variables written in default syntax.
		$template = self::replaceVariables($template);

		// Replace the variables written in alternative syntax.
		$template = self::replaceAlternativeVariables($template);

		// Output the document.
		$template->output('docx', 'Invoice A19583.docx');
	}


	/**
	 * Example of outputting a document in PDF format.
	 *
	 * @static
	 *
	 * @return  void
	 *
	 */
	public static function outputPDF(): void
	{
		$template = self::getTemplate();
		$template = self::setDocumentInfo($template);
		$template = self::replaceVariables($template);
		$template = self::replaceAlternativeVariables($template);

		$template->output('pdf', 'Invoice A19583.pdf');
	}


	/**
	 * Example of saving a document in different formats.
	 *
	 * @param   string   $format
	 * @static
	 *
	 * @return  void
	 *
	 */
	public static function saveDocument(string $format = 'docx'): void
	{
		$template = self::getTemplate();
		$template = self::replaceVariables($template);
		$template = self::replaceAlternativeVariables($template);
		$template = self::setDocumentInfo($template);

		// Use the full path outside the working directory.
		$template->save($_SERVER['DOCUMENT_ROOT'] . "/Invoice A19583.$format", $format);
	}


	/**
	 * Example of outputting the first document section in different formats.
	 *
	 * @param   string   $format
	 * @static
	 *
	 * @return  void
	 *
	 */
	public static function outputFirstSection(string $format = 'docx'): void
	{
		$template = self::getTemplate();
		$template = self::setDocumentInfo($template);

		// Truncate the document to the first section.
		// If you have images in removed sections, you need to remove them manually to reduce the document size.
		// You can also copy the working section after truncating, the first only section.
		$template->useSection(1);

		// Replace the variables AFTER copying, as we do not need to process deleted content.
		$template = self::replaceVariables($template);
		$template->output($format, "Invoice A19583.$format");
	}


	/**
	 * Example of copying the whole document and outputting it in different formats.
	 *
	 * @param   string   $format
	 * @static
	 *
	 * @return  void
	 *
	 */
	public static function copyDocument(string $format = 'docx'): void
	{
		$template = self::getTemplate();
		$template = self::setDocumentInfo($template);

		// Replace the variables BEFORE copying. It may be faster to replace AFTER copying, which depends on your document.
		$template = self::replaceVariables($template);
		$template = self::replaceAlternativeVariables($template);

		// Make one copy of the whole document
		$template->repeat(1);
		$template->output($format, "Invoice A19583.$format");
	}


	/**
	 * Example of copying the first document section and outputting it in different formats.
	 *
	 * @param   string   $format
	 * @static
	 *
	 * @return  void
	 *
	 */
	public static function copyFirstSection(string $format = 'docx'): void
	{
		$template = self::getTemplate();
		$template = self::setDocumentInfo($template);

		// Make one copy of the first section and append it to the end of the document.
		$template->repeat(1, 1);

		// Replace the variables AFTER copying. It may be faster to replace BEFORE copying, which depends on your document.
		$template = self::replaceVariables($template);
		$template = self::replaceAlternativeVariables($template);

		$template->output($format, "Invoice A19583.$format");
	}


	/**
	 * Example of deleting an image and outputting the document in different formats.
	 *
	 * @param   string   $format
	 * @static
	 *
	 * @return  void
	 *
	 */
	public static function deleteSignature(string $format = 'docx'): void
	{
		$template = self::getTemplate();
		$template = self::setDocumentInfo($template);
		$template = self::replaceVariables($template);
		$template = self::replaceAlternativeVariables($template);

		// Refer to an image id according to the Word enumerator.
		$template->deleteImage(2);

		$template->output($format, "Invoice A19583.$format");
	}


	/**
	 * Example of replacing an image with another image.
	 *
	 * @param   string   $format
	 * @static
	 *
	 * @return  void
	 *
	 */
	public static function replaceSignature(string $format = 'docx'): void
	{
		$template = self::getTemplate();
		$template = self::setDocumentInfo($template);
		$template = self::replaceVariables($template);
		$template = self::replaceAlternativeVariables($template);

		// The path of a new image.
		$new_img_path = dirname(__FILE__) . '/img_samples/xx_signature.png';

		// Refer to an image id according to the Word enumerator.
		$template->replaceImage(2, $new_img_path);

		$template->output($format, "Invoice A19583.$format");
	}


	/**
	 * Example of outputting a document in HTML/XHTML format.
	 *
	 * @param   bool  $xhtml
	 * @static
	 *
	 * @return  void
	 *
	 */
	public static function outputHTML(bool $xhtml = false): void
	{
		$template = self::getTemplate();
		$template = self::setDocumentInfo($template);

		$template->useSection(1);
		$template = self::replaceVariables($template);

		// Set any styles and scripts.
		$template->embedStyleSheet('table:first-of-type tr:first-of-type td:last-child span {color: red}');
		$template->embedStyleSheet('.Table4_E10 {position: relative} .P7 {position: absolute; bottom: -14px}');
		$template->embedScript("console.log('Hello World!');");

		$template->output($xhtml ? 'xhtml' : 'html');
	}


	/**
	 * Example of creating the HTML document adapted to email and mailing it to a reciever.
	 *
	 * @param   string   $from   Sender email.
	 * @param   string   $to     Reciever email.
	 * @static
	 *
	 * @return  void
	 *
	 */
	public static function mailHTML(string $from, string $to): void
	{
		$template = self::getTemplate();
		$template->useSection(1);
		$template = self::replaceVariables($template);

		// Set some simple styles like .class, #id, table.class, p#id, *, div, span, p, table... other elements.
		$template->embedStyleSheet('.T1 {color: purple}');

		// Use a temporary working directory, that is removed by the destructor.
		// As we do not need this file after sending an email.
		$path = $template->save('invoice.html', 'mail');

		$contents = file_get_contents($path);

		$imageDirUrl = 'http://' . $_SERVER['HTTP_HOST'] . str_replace('//', '/', str_replace($_SERVER['DOCUMENT_ROOT'], '/', dirname(__FILE__))) . '/img_samples/';

		// Set images instead of placeholders. We trim the document images, because they can be invalid for the mail. We replace them with a HTML comment.
		// You can also use the template variables and replace them with HTML elements, but AFTER generating the HTML document.
		$contents = str_ireplace('<!--[image_1]-->', '<img src="' . $imageDirUrl . 'logo.png" />', $contents);
		$contents = str_ireplace('<!--[image_2]-->', '<img src="' . $imageDirUrl . 'jd_signature.png" style="width: 107px; height: 32px;" widht="107" height="32" />', $contents);

		$headers = [
			'MIME-Version: 1.0',
			'Content-type: text/html; charset=utf-8',
			'To: Customer <' . $to . '>',
			'From: Your company <' . $from . '>',
		];

		// Mail example function.
		mail($to, 'Invoice A19583', $contents, implode("\r\n", $headers));
	}


	/**
	 * Trigger an error to process it with a custom error handler.
	 *
	 * @static
	 *
	 * @return  void
	 *
	 */
	public static function simulateError(): void
	{
		$template = self::getTemplate();

		// Use a not existing section to trigger an error.
		$template->useSection(100500);
	}


	/**
	 * Customer error handler example.
	 *
	 * @param    string   $errorText
	 * @param    string   $errorStatus
	 * @static
	 *
	 * @return  void
	 *
	 */
	public static function errorHandler(string $errorText, string $errorStatus): void
	{
		echo sprintf('Error processing using custom error handler. Error status: <b>%s</b>, error text: <b>%s</b>', $errorStatus, $errorText);
		exit;
	}
}
