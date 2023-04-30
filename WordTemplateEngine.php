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
 * @version    1.0.3
 * @author     Philip Sorokin <philip.sorokin@gmail.com>
 * @copyright  2021 - Philip Sorokin
 * @license    http://www.gnu.org/licenses/gpl-3.0.html GNU General Public License
 * @note       This program is distributed in the hope that it will be useful - WITHOUT
 * ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
 */

class WordTemplateEngine
{
	/**
	 * An alternative error handler.
	 *
	 * @var   callable
	 */
	protected $errorHandler = null;

	/**
	 * Main document namespace.
	 *
	 * @var   string
	 */
	protected $namespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

	/**
	 * User defined embedded style sheets.
	 *
	 * @var   array
	 */
	protected $embeddedStyleSheets = [];

	/**
	 * User defined embedded scripts.
	 *
	 * @var   array
	 */
	protected $embeddedScripts = [];

	/**
	 * User defined external style sheets.
	 *
	 * @var   array
	 */
	protected $styleSheets = [];

	/**
	 * User defined external scripts.
	 *
	 * @var   array
	 */
	protected $scripts = [];

	/**
	 * Locale for conversion.
	 *
	 * @var   string
	 */
	protected $locale = 'C.UTF-8';

	/**
	 * Conversion filter.
	 *
	 * @var   string
	 */
	protected $outputFilter = null;

	/**
	 * Send output headers.
	 *
	 * @var   bool
	 */
	protected $sendOutputHeaders = true;

	/**
	 * Temporary directory path.
	 *
	 * @var   string
	 */
	protected $tmpDir = null;

	/**
	 * Working directory path.
	 *
	 * @var   string
	 */
	protected $docDir = null;

	/**
	 * The document.
	 *
	 * @var   DOMDocument
	 */
	protected $document = null;

	/**
	 * The document headers.
	 *
	 * @var   array   An array with DOMDocument objects.
	 */
	protected $headers = [];

	/**
	 * The document footers.
	 *
	 * @var   array   An array with DOMDocument objects.
	 */
	protected $footers = [];

	/**
	 * The document rels.
	 *
	 * @var   array   An array with DOMDocument objects.
	 */
	protected $rels = [];

	/**
	 * Document rels elements.
	 *
	 * @var   array   An array with DOMElement objects.
	 */
	protected $relsElements = [];

	/**
	 * Document rels targets.
	 *
	 * @var   array   A multidimensional array with DOMElement objects.
	 */
	protected $relsTargets = [];

	/**
	 * The document core properties.
	 *
	 * @var   DOMDocument
	 */
	protected $props = null;

	/**
	 * The document extended properties.
	 *
	 * @var   DOMDocument
	 */
	protected $appProps = null;

	/**
	 * The document paragraphs.
	 *
	 * @var   array
	 */
	protected $paragraphs = null;

	/**
	 * The document table rows.
	 *
	 * @var   array
	 */
	protected $rows = null;

	/**
	 * The document images indexed by blip IDs.
	 *
	 * @var   array
	 */
	protected $images = null;

	/**
	 * Image blips map.
	 *
	 * @var   array
	 */
	protected $blipsMap = [];

	/**
	 * The anchor character of a variable.
	 *
	 * @var   string
	 */
	protected $anchor_char = '$';

	/**
	 * The open character of a variable.
	 *
	 * @var   string
	 */
	protected $open_char = '{';

	/**
	 * The close character of a variable.
	 *
	 * @var   string
	 */
	protected $close_char = '}';


	/**
	 * Constructor.
	 *
	 * @param    string         $template        The full path to the .docx template.
	 * @param    string|null    $tmp_path        Optional. The full path to the temporary directory where the working files are created. By default, the site root or a current directory is used.
	 * @param    callable       $errorHandler    Optional. A function to replace the default error handler. The callback takes two string parameters: the error description and a unique error status.
	 *
	 */
	public function __construct(string $template, ?string $tmp_path = null, callable $errorHandler = null)
	{
		$tmp_path = $tmp_path ?? $_SERVER['DOCUMENT_ROOT'] ?? '';

		if ($errorHandler)
		{
			$this->errorHandler = $errorHandler;
		}

		if (version_compare(PHP_VERSION, '7.1', '<'))
		{
			$this->raiseError("You have to use PHP version 7.1 or higher to run WordTemplateEngine.", 'php_version_error');
		}

		$base = basename($template);

		if (!file_exists($template))
		{
			$this->raiseError("Template '$base' not found.", 'template_not_found');
		}

		$this->tmpDir = implode(DIRECTORY_SEPARATOR, array_filter([$tmp_path, 'temp_wte_' . md5(strval(random_int(0, 1000000000)))])) . DIRECTORY_SEPARATOR;
		mkdir($this->tmpDir);

		$this->docDir = $this->tmpDir . 'doc' . DIRECTORY_SEPARATOR;
		mkdir($this->docDir);

		$bin = $this->docDir . $base;
		copy($template, $bin);

		$zip = new ZipArchive;
		$zip->open($bin);

		if (!$zip->extractTo($this->docDir))
		{
			$this->raiseError("Unable to extract template '$base'.", 'extract_error');
		}

		$zip->close();
		unlink($bin);

		$types = new DOMDocument;

		if (!$types->load($this->docDir . '[Content_Types].xml'))
		{
			$this->raiseError("Unable to process template '$base'.", 'process_error');
		}

		foreach($types->getElementsByTagName('Override') as $override)
		{
			$type = preg_replace('#^.+?\.([^.+]+)[^.]*$#', '$1', $override->getAttribute('ContentType'));
			$path = $this->docDir . preg_replace('#^[/\\\]#', '', $override->getAttribute('PartName'));

			if ($type === 'main')
			{
				$this->document = new DOMDocument;
				$this->document->load($path);
			}
			else if ($type === 'header')
			{
				$doc = new DOMDocument;
				$doc->load($path);

				$this->headers[] = $doc;
			}
			else if ($type === 'footer')
			{
				$doc = new DOMDocument;
				$doc->load($path);

				$this->footers[] = $doc;
			}
			else if ($type === 'core-properties')
			{
				$this->props = new DOMDocument;
				$this->props->load($path);
			}
			else if ($type === 'extended-properties')
			{
				$this->appProps = new DOMDocument;
				$this->appProps->load($path);
			}
		}

		foreach($types->getElementsByTagName('Default') as $default)
		{
			if ($content_type = $default->getAttribute('ContentType'))
			{
				if (mb_stripos($content_type, 'relationships') !== false)
				{
					$relExt = $default->getAttribute('Extension');
					break;
				}
			}
		}

		if (!empty($relExt))
		{
			foreach(array_merge([$this->document], $this->headers, $this->footers) as $item)
			{
				$base = basename($item->baseURI);
				$dir = dirname($item->baseURI);

				if (file_exists($file = $dir . DIRECTORY_SEPARATOR . "_$relExt" . DIRECTORY_SEPARATOR . $base . ".$relExt"))
				{
					$doc = new DOMDocument();
					$doc->load($file);

					$this->rels[$file] = $doc;

					$this->relsElements[$base] = [];
					$relsElements = $doc->getElementsByTagName('Relationship');

					foreach($relsElements as $element)
					{
						$id = $element->getAttribute('Id');
						$target = $element->getAttribute('Target');

						$this->relsElements[$base][$id] = $element;

						if ($target && (mb_strpos($target, '${') !== false || mb_strpos($target, '~(') !== false) && preg_match_all('#[~$][{(](.+?)[)}]#', $target, $matches))
						{
							foreach($matches[1] as $match)
							{
								$varName = mb_strtolower($match);

								if (empty($this->relsTargets[$varName]))
								{
									$this->relsTargets[$varName] = [];
								}

								$this->relsTargets[$varName][] = $element;
							}
						}
					}
				}
			}
		}
	}


	/**
	 * Destructor. Calls the cleaner to remove the working directory with temporary files.
	 *
	 */
	public function __destruct()
	{
		$this->clean();
	}


	/**
	 * The locale that is set before conversion the .docx template into other formats. Default is 'C.UTF-8'.
	 *
	 * @param   string   $locale   UNIX locale.
	 *
	 * @return  void
	 *
	 */
	public function setLocale(string $locale): void
	{
		$this->locale = $locale;
	}


	/**
	 * If you need to overwrite a default conversion filter, define it before saving / outputting the document.
	 *
	 * @param   string   $filter   LibreOffice filter parameter, e.g. 'HTML:EmbedImages'.
	 *
	 * @return  void
	 *
	 */
	public function setOutputFilter(string $filter): void
	{
		$this->outputFilter = $filter;
	}


	/**
	 * Call this method with the FALSE parameter before outputting the document if you want to send your own HTTP headers.
	 *
	 * @param   bool   $bool
	 *
	 * @return  void
	 *
	 */
	public function sendOutputHeaders(bool $bool): void
	{
		$this->sendOutputHeaders = $bool;
	}


	/**
	 * Method to embed a custom stylesheet into the HEAD section of the document that is to be converted into HTML or XHTML format.
	 *
	 * @param   string   $stylesheet
	 *
	 * @return  void
	 *
	 */
	public function embedStyleSheet(string $stylesheet): void
	{
		$this->embeddedStyleSheets[] = $stylesheet;
	}


	/**
	 * Method to embed a custom script into the HEAD section of a document that is to be converted into HTML or XHTML format.
	 *
	 * @param   string   $script
	 *
	 * @return  void
	 *
	 */
	public function embedScript(string $script): void
	{
		$this->embeddedScripts[] = $script;
	}


	/**
	 * Method to append an external stylesheet into the HEAD section of a document that is to be converted into HTML or XHTML format.
	 *
	 * @param   string   $url
	 *
	 * @return  void
	 *
	 */
	public function addStyleSheet(string $url): void
	{
		$this->styleSheets[] = $url;
	}


	/**
	 * Method to append an external script into the HEAD section of a document that is to be converted into HTML or XHTML format.
	 *
	 * @param   string   $url
	 *
	 * @return  void
	 *
	 */
	public function addScript(string $url): void
	{
		$this->scripts[] = $url;
	}


	/**
	 * Set extended document metadata - Company.
	 *
	 * @param   string   $companyName
	 *
	 * @return  void
	 *
	 */
	public function setCompany(string $companyName): void
	{
		$this->setAppData('Company', $companyName);
	}


	/**
	 * Set extended document metadata - Manager.
	 *
	 * @param   string   $manager
	 *
	 * @return  void
	 *
	 */
	public function setManager(string $manager): void
	{
		$this->setAppData('Manager', $manager);
	}


	/**
	 * Set core document metadata - title. It also displays as the page title in documents opened in a browser: PDF, HTML, XHTML.
	 *
	 * @param   string   $title
	 *
	 * @return  void
	 *
	 */
	public function setTitle(string $title): void
	{
		$this->setMetaData('title', $title);
	}


	/**
	 * Set core document metadata - creator. Defines the document creator, clears lastModifiedBy properties.
	 *
	 * @param   string   $author
	 *
	 * @return  void
	 *
	 */
	public function setAuthor(string $author): void
	{
		foreach(['creator', 'lastModifiedBy'] as $key => $name)
		{
			$this->setMetaData($name, !$key ? $author : null);
		}
	}


	/**
	 * Set core document metadata - created. Defines the document creation DateTime, clears modified and lastPrinted DateTime properties.
	 *
	 * @param   string|null   $time   Optional. DateTime in ISO 8601 format. If undefined or null current DateTime is used.
	 *
	 * @return  void
	 *
	 */
	public function setTime(?string $time = null): void
	{
		if (is_null($time))
		{
			$time = new DateTime;
			$time = $time->format('c');
		}

		foreach(['created', 'modified', 'lastPrinted'] as $key => $name)
		{
			$this->setMetaData($name, !$key ? $time : null);
		}
	}


	/**
	 * Set core document metadata - subject.
	 *
	 * @param   string   $subject
	 *
	 * @return  void
	 *
	 */
	public function setSubject(string $subject): void
	{
		$this->setMetaData('subject', $subject);
	}


	/**
	 * Set core document metadata - keywords.
	 *
	 * @param   string   $keywords
	 *
	 * @return  void
	 *
	 */
	public function setKeywords(string $keywords): void
	{
		$this->setMetaData('keywords', $keywords);
	}


	/**
	 * Set core document metadata - description.
	 *
	 * @param   string   $description
	 *
	 * @return  void
	 *
	 */
	public function setDescription(string $description): void
	{
		$this->setMetaData('description', $description);
	}


	/**
	 * Set core document metadata - category.
	 *
	 * @param   string   $category
	 *
	 * @return  void
	 *
	 */
	public function setCategory(string $category): void
	{
		$this->setMetaData('category', $category);
	}


	/**
	 * Set core document metadata - content status.
	 *
	 * @param   string   $contentStatus
	 *
	 * @return  void
	 *
	 */
	public function setStatus(string $contentStatus): void
	{
		$this->setMetaData('contentStatus', $contentStatus);
	}


	/**
	 * Drop all core document properties and two extended properties: Company, Manager.
	 * After calling this method you have to redefine the document creation time, the title and the document creator; otherwise, the document will be broken.
	 *
	 * @return  void
	 *
	 */
	public function dropMetaData(): void
	{
		$root = $this->props->documentElement;
		$cnt = $root->childNodes->length;

		while($cnt-- > 0)
		{
			$root->removeChild($root->childNodes[$cnt]);
		}

		$this->setAppData('Company', null);
		$this->setAppData('Manager', null);
	}


	/**
	 * Remove all the sections in the template except the section with an index passed as the argument. This method truncates the document to the certain section.
	 * Note! It must be called before any replacements, because the document paragraphs and rows are cached for better performance.
	 * Note! If you use page breaks inside a section, the next page must begin with a paragraph. Otherwise, the document can be broken.
	 *
	 * @param   int   $idx   The section index.
	 *
	 * @return  void
	 *
	 */
	public function useSection(int $idx): void
	{
		$sections = $this->markSections();

		if (!isset($sections[$idx - 1]))
		{
			$this->raiseError("Section $idx does not exist.", 'section_not_found');
		}

		$body = $this->document->getElementsByTagNameNS($this->namespace, 'body')->item(0);

		$cnt = $body->childNodes->length;
		$working = $idx;

		while($cnt-- > 0)
		{
			$element = $body->childNodes->item($cnt);
			$secIdx = $element->getAttribute('wte_section');

			if ($secIdx)
			{
				$working = $secIdx;
			}

			$working != $idx ? $body->removeChild($element) : $element->removeAttribute('wte_section');
		}

		$section = $sections->item(0);

		if ($section->parentNode->tagName !== $body->tagName)
		{
			$paragraph = $section->parentNode->parentNode;

			$body->appendChild($section);
			$body->removeChild($paragraph);
		}
	}


	/**
	 * Copy the document one or several times. If the second argument is passed, it creates only a section copy and appends it to the end of the document.
	 * Note! It must be called before or after any replacements, because the document paragraphs and rows are cached for better performance. The call order depends on your document.
	 *
	 * @param   int   $cnt   Optional. How many copies are created. Default 1.
	 * @param   int   $idx   Optional. Repeat only one section.
	 *
	 * @return  void
	 *
	 */
	public function repeat(int $cnt = 1, int $idx = null): void
	{
		if ($cnt < 1)
		{
			return;
		}

		$body = $this->document->getElementsByTagNameNS($this->namespace, 'body')->item(0);

		$clones = [];

		if (!$idx)
		{
			foreach($body->childNodes as $node)
			{
				$clones[] = $node->cloneNode(true);
			}
		}
		else
		{
			$sections = $this->markSections();

			if (!isset($sections[$idx - 1]))
			{
				$this->raiseError("Section $idx does not exist.", 'section_not_found');
			}

			if ($idx - 2 < 0)
			{
				$node = $body->childNodes->item(0);
			}
			else
			{
				$node = $sections->item($idx - 2)->parentNode->parentNode->nextSibling;
			}

			$node = $idx - 2 < 0 ? $body->childNodes->item(0) : $sections->item($idx - 2)->parentNode->parentNode->nextSibling;

			while($node)
			{
				$clones[] = $node->cloneNode(true);

				if ($node->getAttribute('wte_section'))
				{
					break;
				}

				$node = $node->nextSibling;
			}
		}

		while($cnt-- > 0)
		{
			$node = $body->childNodes[$body->childNodes->length - 1];
			$separator = $node->cloneNode(true);
			$body->removeChild($node);

			$pPrs = $body->getElementsByTagNameNS($this->namespace, 'pPr');

			if ($pPrs->length)
			{
				$pPr = $pPrs->item($pPrs->length - 1);

				$paragraph = $pPr->parentNode->cloneNode();
				$paragraph->appendChild($pPr->cloneNode());
				$paragraph->childNodes->item(0)->appendChild($separator->cloneNode(true));

				$body->appendChild($paragraph);

				foreach($clones as $clone)
				{
					$body->appendChild($clone->cloneNode(true));
				}

				if ($idx)
				{
					$body->appendChild($sections->item($sections->length - 1));
				}
			}
		}
	}


	/**
	 * Method to replace a template variable with the $replacement value.
	 * A template variable must have the following format: ${name}, where 'name' is passed as the $varName argument.
	 *
	 * @param   string        $varName       The variable name.
	 * @param   string|null   $replacement   The replacement.
	 *
	 * @return  void
	 *
	 */
	public function setValue(string $varName, ?string $replacement): void
	{
		$varName = mb_strtolower($varName);
		$paragraphs = $this->getParagraphs($varName);

		$search = $this->anchor_char . $this->open_char . $varName . $this->close_char;
		$replacement = (string) $replacement;

		$searchLen = mb_strlen($search);

		foreach($paragraphs as $paragraph)
		{
			$contentNode = null;

			while (mb_stripos($paragraph->textContent, $search) !== false)
			{
				$working = [];
				$processed = 0;

				$texts = $this->getElements($paragraph, 'w:t');
				$textContent = '';

				foreach($texts as $text)
				{
					$textContent .= $text->textContent;
				}

				$pos = mb_stripos($textContent, $search);

				foreach($texts as $text)
				{
					$processed += mb_strlen($text->textContent);

					if ($processed <= $pos)
					{
						continue;
					}

					$working[] = $text;

					if ($processed >= $pos + $searchLen)
					{	
						break;
					}
				}

				$contentNode = array_shift($working);

				if (!$contentNode)
				{
					break;
				}

				$textContent = $contentNode->textContent;

				foreach($working as $key => $text)
				{
					$textContent .= $text->textContent;
					$text->textContent = '';
				}

				$contentNode->textContent = str_ireplace($search, $replacement, $textContent, $cnt);

				if (!$cnt)
				{
					break;
				}
			}
		}

		if (isset($this->relsTargets[$varName]))
		{
			foreach($this->relsTargets[$varName] as $element)
			{
				$element->setAttribute('Target', str_ireplace($search, $replacement, $element->getAttribute('Target')));
			}
		}
	}


	/**
	 * Method to clone a table row with an anchor. An integer counter followed by '#' is appended to the names of all variables inside the row including the anchor.
	 * E.g. you have an anchor variable: ${key} and other variables inside the row: ${customer}, ${order_id}, they are replaced with: ${key#1}, ${customer#1}, ${order_id#1}; ${key#2}, ${customer#2}, ${order_id#2} etc.
	 *
	 * @param   string   $varName   The anchor variable.
	 * @param   int      $cnt       How many rows you need.
	 *
	 * @return  void
	 *
	 */
	public function cloneRow(string $varName, int $cnt): void
	{
		$varName = mb_strtolower($varName);
		$rows = $this->getRows($varName);

		foreach($rows as $key => $row)
		{
			$search = $this->anchor_char . $this->open_char . $varName . $this->close_char;
			$searchLen = mb_strlen($search);

			if (mb_stripos($row->textContent, $search) !== false)
			{
				$i = 0;

				while($i++ < $cnt)
				{
					$clone = $row->cloneNode(true);
					$this->replaceRowVariables($clone, $i);
					$row->parentNode->insertBefore($clone, $row);
				}

				$row->parentNode->removeChild($row);
			}
		}

		unset($this->rows[$varName]);
	}


	/**
	 * Call this method before replacements to use alternative variable syntax ~(name) instead of default ${name}. You can also switch back to default syntax.
	 * Use the alternative syntax to replace the variables inside hyperlinks and other targets.
	 *
	 * @param   bool   $bool   Whether to use alternative variable syntax.
	 *
	 * @return  void
	 *
	 */
	public function alternativeSyntax(bool $bool): void
	{
		$this->anchor_char = $bool ? '~' : '$';
		$this->open_char = $bool ? '(' : '{';
		$this->close_char = $bool ? ')' : '}';
	}


	/**
	 * Replace an image in the document with another image.
	 *
	 * @param   int      $imageID       An image id according to the Word enumerator.
	 * @param   string   $replacement   The path of a new image.
	 *
	 * @return  void
	 *
	 */
	public function replaceImage(int $imageID, string $replacement): void
	{
		if (!file_exists($replacement))
		{
			$this->raiseError("Replacement image does not exist.", 'replace_image_not_exists');
		}

		if ($images = $this->getImages($imageID))
		{
			foreach($images as list($image, $base, $blipID))
			{
				if ($rel = $this->relsElements[$base][$blipID] ?? null)
				{
					$original = $this->docDir . 'word' . DIRECTORY_SEPARATOR . $rel->getAttribute('Target');

					copy($replacement, $original);
				}
			}
		}
	}


	/**
	 * Completely removes images with a certain id, it deletes both the elements and the file.
	 *
	 * @param   int   $imageID   An image id according to the Word enumerator.
	 *
	 * @return  void
	 *
	 */
	public function deleteImage(int $imageID): void
	{
		if ($images = $this->getImages($imageID))
		{
			foreach($images as list($image, $base, $blipID))
			{
				if ($rel = $this->relsElements[$base][$blipID] ?? null)
				{
					$target = $rel->getAttribute('Target');
					$rel->parentNode->removeChild($rel);

					if (file_exists($file = $this->docDir . 'word' . DIRECTORY_SEPARATOR . $target))
					{
						unlink($file);
					}

					unset($this->relsElements[$base][$blipID]);
				}

				$image->parentNode->parentNode->removeChild($image->parentNode);
			}

			unset($this->images[$this->blipsMap[$imageID]]);
			unset($this->blipsMap[$imageID]);
		}
	}


	/**
	 * Creates a document from the processed template.
	 *
	 * @param   string   $destination   The destination path. Can be full or relative, e.g. '/var/www/newdoc.docx', 'newdoc.pdf'. The root of a relative path is the temporary directory.
	 * @param   string   $format        Optional. The document format: 'docx' (default), 'pdf', 'html', 'xhtml', 'mail' (HTML adapted to email).
	 *
	 * @return  string   The full path to the created document.
	 *
	 */
	public function save(string $destination, string $format = 'docx'): string
	{
		$format = strtolower($format);

		if (!in_array($format, ['docx', 'pdf', 'html', 'xhtml', 'mail']))
		{
			$this->raiseError("Unsupported format '$format'.", 'unsupported_format');
		}

		if (empty($destination))
		{
			$this->raiseError("The destination path cannot be empty.", 'empty_destination');
		}

		if (mb_strpos($destination, '/') !== 0 && mb_strpos($destination, '\\') !== 0 && (DIRECTORY_SEPARATOR === '/' || mb_strpos($destination, ':') === false))
		{
			$destination = $this->tmpDir . $destination;
		}

		$this->setMetaData('revision', 1);
		$this->setMetaData('lastPrinted', null);
		$this->setAppData('TotalTime', 0);

		$this->document->save($this->document->baseURI);
		$this->props->save($this->props->baseURI);
		$this->appProps->save($this->appProps->baseURI);

		foreach($this->headers as $header)
		{
			$header->save($header->baseURI);
		}

		foreach($this->footers as $footer)
		{
			$footer->save($footer->baseURI);
		}

		foreach($this->rels as $rel)
		{
			$rel->save($rel->baseURI);
		}

		$tmp = $this->tmpDir . 'output.docx';
		$zip = new ZipArchive();

		$zip->open($tmp, ZIPARCHIVE::CREATE);
		$source = realpath($this->docDir);

		$files = new RecursiveIteratorIterator(new RecursiveDirectoryIterator($this->docDir, RecursiveDirectoryIterator::SKIP_DOTS), RecursiveIteratorIterator::SELF_FIRST);

		foreach ($files as $key => $file)
		{
			if (!$key)
			{
				continue;
			}

			$file = realpath($file);

			if (is_dir($file) === true)
			{
				$zip->addEmptyDir(str_replace($source . DIRECTORY_SEPARATOR, '', $file . DIRECTORY_SEPARATOR));
			}

			else if (is_file($file) === true)
			{
				$zip->addFromString(str_replace($source . DIRECTORY_SEPARATOR, '', $file), file_get_contents($file));
			}
		}

		$zip->close();

		if ($format !== 'docx')
		{
			if (!function_exists('exec'))
			{
				$this->raiseError("You have to enable 'exec' function for conversion from WORD to " . strtoupper($format) . ".", 'exec_function_disabled');
			}

			exec('dpkg-query -l libreoffice', $output, $status);

			if ($status)
			{
				$this->raiseError("You have to install LibreOffice package for conversion from WORD to " . strtoupper($format) . ".", 'libreoffice_not_installed');
			}

			$convert_to = [$format === 'pdf' ? $format : 'html'];

			if ($this->outputFilter)
			{
				$convert_to[] = $this->outputFilter;
			}
			else if ($format !== 'pdf')
			{
				$convert_to[] = 'XHTML Writer File';
			}

			putenv('LC_ALL=' . $this->locale);
			putenv('HOME=/tmp');

			exec('soffice --headless --convert-to ' . escapeshellarg(implode(':', $convert_to)) . ' --outdir ' . escapeshellarg(dirname($tmp)) . ' ' . escapeshellarg($tmp));
			$tmp = $this->tmpDir . "output.{$convert_to[0]}";

			if ($format !== 'pdf')
			{
				$contents = file_get_contents($tmp);
				$assertions = [];

				foreach($this->styleSheets as $url)
				{
					$assertions[] = '<link rel="stylesheet" type="text/css" href="' . $url . '" />';
				}

				foreach($this->scripts as $url)
				{
					$assertions[] = '<script type="text/javascript" src="' . $url . '"></script>';
				}

				foreach($this->embeddedStyleSheets as $text)
				{
					$assertions[] = '<style type="text/css">' . $text . '</style>';
				}

				foreach($this->embeddedScripts as $text)
				{
					$assertions[] = '<script type="text/javascript">' . $text . '</script>';
				}

				if (!empty($assertions))
				{
					$contents = preg_replace('#</head\s*>#i', implode('', $assertions) . '$0', $contents, 1);
				}

				$contents = preg_replace('#<[^>]+?\K\b(?:xml:)?(?:lang)\s*=\s*(["\'])[^>]*?\\1\s*#i', '', $contents);
				$contents = preg_replace('#<meta[^>]+\bDCTERMS\.language[^>]*>#i', '', $contents);
				$contents = preg_replace('#[^\s[:print:]]#u', '', $contents);

				if ($format !== 'xhtml')
				{
					if (stripos($contents, '<?xml') !== false)
					{
						$contents = preg_replace('#<\?xml.+?(<html)#is', "<!DOCTYPE HTML>\n$1", $contents, 1);
						$contents = preg_replace('#<[^>]+\K\b(?:xmlns)\s*=\s*(["\'])[^>]*?\\1\s*#i', '', $contents);
						$contents = preg_replace('#<[^>]+\K\bxml:#i', '', $contents);
						$contents = preg_replace('#<meta[^>]+\bcontent\-type[^>]+>#i', '<meta name="content-type" content="text/html" />', $contents);
						$contents = preg_replace('#<meta[^>]+\bcharset[^>]*>#i', '', $contents);
						$contents = preg_replace('#<head[^>]*>\K#i', '<meta charset="utf8" />', $contents);
						$contents = preg_replace('#(<h\b)([^>]*>.*?</h)(\b\s*>)#i', '${1}1${2}1${3}', $contents);
					}

					if ($format === 'mail')
					{
						$contents = preg_replace('#<!--.*?-->#s', '', $contents);

						$mail_image_num = 0;

						$contents = preg_replace_callback('#<img[^>]+\bsrc\s*=\s*[\'"]data:[^>]+>#i', function($m) use (&$mail_image_num) {

							$mail_image_num++;

							return "<!--[image_$mail_image_num]-->";

						}, $contents);

						$contents = preg_replace('#<(link|meta)(?![^>]+(content\-type|charset))[^>]*>#is', '', $contents);
						$contents = preg_replace('#<[^>]+\K\b(?:profile)\s*=\s*(["\'])[^>]*?\\1\s*#i', '', $contents);

						$contents = $this->inlineCss($contents);
					}
				}

				file_put_contents($tmp, $contents);
			}
		}

		rename($tmp, $destination);
		return $destination;
	}


	/**
	 * Creates a document from the processed template and outputs it to the client browser.
	 *
	 * @param   string        $format          Optional. The document format: 'docx' (default), 'pdf', 'html', 'xhtml', 'mail' (HTML adapted to email).
	 * @param   string|null   $fileName        Optional. Defines the filename of downloaded file.
	 * @param   string        $isAttachment    Optional. Forces docx and pdf documents to the certain disposition type: inline or attachment. By default, docx documents are attachments, PDFs are inlines.
	 *
	 * @return  void
	 *
	 */
	public function output(string $format = 'docx', ?string $fileName = null, bool $isAttachment = null): void
	{
		$format = mb_strtolower($format);
		$source = $this->save("output.$format", $format);

		if ($this->sendOutputHeaders)
		{
			$headers = ['Cache-Control: no-store, no-cache, must-revalidate, max-age=0, no-transform'];

			if (in_array($format, ['docx', 'pdf']))
			{
				$headers[] = 'Content-type: application/' . ($format === 'docx' ? 'vnd.openxmlformats-officedocument.wordprocessingml.document' : 'pdf');
				$headers[] = 'Content-Disposition: ' . (!empty($isAttachment) || !isset($isAttachment) && $format === 'docx' ? 'attachment' : 'inline');
				$headers[] = 'Content-Description: File Transfer';
				$headers[] = 'Content-Transfer-Encoding: binary';

				if ($fileName)
				{
					$headers[2] .= '; filename="' . $fileName . '";' . "filename*=utf-8''" . rawurlencode($fileName) . ';';
				}
			}
			else
			{
				$headers[] = 'Content-type: text/html; charset=utf-8';
			}

			array_walk($headers, 'header');
		}

		readfile($source);
	}


	/**
	 * The image getter.
	 *
	 * @param   id   $imageID   An image id according to the Word enumerator.
	 *
	 * @return  array|null
	 *
	 */
	protected function getImages(int $imageID): ?array
	{
		if (!isset($this->images))
		{
			foreach(array_merge([$this->document], $this->headers, $this->footers) as $doc)
			{
				$base = basename($doc->baseURI);

				foreach($doc->getElementsByTagName('drawing') as $image)
				{
					$id = $image->getElementsByTagName('docPr')->item(0)->getAttribute('id');
					$blipID = $image->getElementsByTagName('blip')->item(0)->getAttribute('r:embed');

					$mapKey = "$blipID:$base";

					$this->blipsMap[$id] = $mapKey;

					if (empty($this->images[$mapKey]))
					{
						$this->images[$mapKey] = [];
					}

					$this->images[$mapKey][] = [$image, $base, $blipID];
				}
			}
		}

		return isset($this->blipsMap[$imageID], $this->images[$this->blipsMap[$imageID]]) ? $this->images[$this->blipsMap[$imageID]] : null;
	}


	/**
	 * Get section elements and mark them with a custom attribute 'wte_section'.
	 *
	 * @return  DOMNodeList  List of section elements.
	 *
	 */
	protected function markSections(): DOMNodeList
	{
		$sections = $this->document->getElementsByTagNameNS($this->namespace, 'sectPr');

		foreach($sections as $key => $section)
		{
			$rootElement = $section->parentNode->tagName === 'w:body' ? $section : $section->parentNode->parentNode;
			$rootElement->setAttribute('wte_section', $key + 1);
		}

		return $sections;
	}


	/**
	 * Remove temporary files in the temporary directory.
	 *
	 * @return  void
	 *
	 */
	protected function clean(): void
	{
		if (is_dir($this->tmpDir))
		{
			$files = new RecursiveIteratorIterator(new RecursiveDirectoryIterator($this->tmpDir, RecursiveDirectoryIterator::SKIP_DOTS), RecursiveIteratorIterator::CHILD_FIRST);

			foreach($files as $fileInfo)
			{
				$fileName = $fileInfo->getRealPath();
				$fileInfo->isDir() ? rmdir($fileName) : unlink($fileName);
			}

			rmdir($this->tmpDir);
		}
	}


	/**
	 * Set core document metadata.
	 *
	 * @param   string        $name    Property name.
	 * @param   string|null   $value   Property value. If null is passed, it removes the element.
	 *
	 * @return  void
	 *
	 */
	protected function setMetaData(string $name, ?string $value): void
	{
		$nodes = $this->props->getElementsByTagName($name);

		if (!isset($value))
		{
			$cnt = $nodes->length;

			while($cnt-- > 0)
			{
				$nodes[$cnt]->parentNode->removeChild($nodes[$cnt]);
			}
		}
		else
		{
			if ($nodes->length)
			{
				$this->setMetaData($name, null);
			}

			if (in_array($name, ['title', 'subject', 'creator', 'description']))
			{
				$namespace = 'http://purl.org/dc/elements/1.1/';
			}
			else if (in_array($name, ['created', 'modified']))
			{
				$namespace = 'http://purl.org/dc/terms/';
			}
			else if (in_array($name, ['keywords', 'lastModifiedBy', 'revision', 'lastPrinted', 'category', 'contentStatus']))
			{
				$namespace = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
			}

			if (isset($namespace))
			{
				$node = $this->props->createElementNS($namespace, $name);
				$node->appendChild($this->props->createTextNode($value));

				if (in_array($name, ['created', 'modified']))
				{
					$attr = $this->props->createAttributeNS('http://www.w3.org/2001/XMLSchema-instance', 'type');
					$attr->value = 'dcterms:W3CDTF';

					$node->appendChild($attr);
				}

				$this->props->documentElement->appendChild($node);
			}
		}
	}


	/**
	 * Set extended document metadata.
	 *
	 * @param   string        $name    Property name.
	 * @param   string|null   $value   Property value. If null is passed, it removes the element.
	 *
	 * @return  void
	 *
	 */
	protected function setAppData(string $name, ?string $value): void
	{
		$nodes = $this->appProps->getElementsByTagName($name);

		if (!isset($value))
		{
			$cnt = $nodes->length;

			while($cnt-- > 0)
			{
				$nodes[$cnt]->parentNode->removeChild($nodes[$cnt]);
			}
		}
		else
		{
			if ($nodes->length)
			{
				$this->setAppData($name, null);
			}

			$node = $this->appProps->createElement($name);
			$node->appendChild($this->appProps->createTextNode($value));

			$this->appProps->documentElement->appendChild($node);
		}
	}


	/**
	 * Super fast replacer of native DOM getElementsByTagName method.
	 *
	 * @param    DOMNode   $parent   Parent node.
	 * @param    string       $tagName  Tag name.
	 *
	 * @return   array   List of elements.
	 *
	 */
	protected function getElements(DOMNode $parent, string $tagName): array
	{
		$elements = [];

		if ($element = $parent->firstChild)
		{
			do {

				if (isset($element->tagName) && ($tagName === '*' || $element->tagName === $tagName))
				{
					$elements[] = $element;
				}

				if ($element->childNodes->length && ($tagName === '*' || $element->tagName !== $tagName))
				{					
					$elements = array_merge($elements, $this->getElements($element, $tagName));
				}

			} while ($element = $element->nextSibling);
		}

		return $elements;
	}


	/**
	 * Iterate the document to fetch indexed arrays with rows and paragraphs containing variables.
	 *
	 * @param    int     $mode  The parser mode. 1 is paragraphs selector, 2 is rows selector.
	 *
	 * @return   array   Cached list of elements.
	 *
	 */
	protected function parse(int $mode)
	{
		$result = [];

		foreach(array_merge([$this->document], $this->headers, $this->footers) as $key => $dom)
		{
			$body = $key ? $dom->documentElement : $dom->documentElement->childNodes->item(0);

			if ($element = $body->firstChild)
			{
				do {

					if (mb_strpos($element->textContent, '${') !== false || mb_strpos($element->textContent, '~(') !== false)
					{
						$elements = [];

						if ($element->nodeName === 'w:p')
						{
							$elements[] = $element;
						}
						else
						{
							$elements = $this->getElements($element, 'w:p');
						}

						foreach($elements as $item)
						{
							if ($mode === 1 || $item->parentNode && $item->parentNode->parentNode && $item->parentNode->parentNode->nodeName === 'w:tr')
							{
								if (preg_match_all('#[~$][{(](.+?)[)}]#', $item->textContent, $matches))
								{
									foreach($matches[1] as $name)
									{
										$name = mb_strtolower($name);

										if ($mode === 1)
										{
											if (empty($this->paragraphs[$name]))
											{
												$this->paragraphs[$name] = [];
											}

											$result[$name][] = $item;
										}
										else
										{
											if (empty($result[$name]))
											{
												$result[$name] = [];
											}

											$result[$name][] = $item->parentNode->parentNode;
										}
									}
								}
							}
						}
					}

				} while ($element = $element->nextSibling);
			}
		}

		return $result;
	}


	/**
	 * The paragraphs getter.
	 *
	 * @param    string   $varName   A variable name.
	 *
	 * @return   array    Cached list of paragraph elements.
	 *
	 */
	protected function getParagraphs(string $varName): array
	{
		if (!isset($this->paragraphs))
		{
			$this->paragraphs = $this->parse(1);
		}

		return $this->paragraphs[$varName] ?? [];
	}


	/**
	 * The rows getter.
	 *
	 * @param    string   $varName   A variable name.
	 *
	 * @return   array    Cached list of row elements.
	 *
	 */
	protected function getRows(string $varName): array
	{
		if (!isset($this->rows))
		{
			$this->rows = $this->parse(2);
		}

		return $this->rows[$varName] ?? [];
	}


	/**
	 * Replace row variables with the same variables containing a row number.
	 *
	 * @param   DOMElement    $row   A row element.
	 * @param   int           $num   A row number.
	 *
	 * @return  void
	 *
	 */
	protected function replaceRowVariables(DOMElement $row, int $num): void
	{
		$texts = $this->getElements($row, 'w:t');

		$start = null;
		$open = null;
		$close = null;

		foreach($texts as $text)
		{
			if (mb_strpos($text->textContent, $this->anchor_char) !== false)
			{
				$start = true;
				$close = null;
			}

			$openPos = mb_strpos($text->textContent, $this->open_char);

			if ($start && !$close && $openPos !== false)
			{
				$open = true;
			}

			$closePos = mb_strpos($text->textContent, $this->close_char);

			if ($closePos !== false && $openPos !== $closePos)
			{
				$close = true;

				if ($start && $open)
				{
					$text->textContent = mb_substr($text->textContent, 0, $closePos) . "#$num" . mb_substr($text->textContent, $closePos);
				}

				$start = null;
				$open = null;
			}
		}
	}


	/**
	 * Inline simple CSS styles, append the rules to the STYLE attribute of elements. This is not a comprehensive CSS parser, it's intended for adapting HTML to email.
	 * The method can process the following styles: .class, #id, table.class, p#id, *, div, span, p, table... other elements.
	 * It cannot process nested selectors like .parent .child or #parent > .child, media queries etc.
	 *
	 * @param   string   $contents   HTML document.
	 *
	 * @return  void
	 *
	 */
	protected function inlineCss(string $contents): string
	{
		$styles = [];

		$contents = preg_replace_callback('#<style[^>]*>\s*(.*?)\s*</style\s*>#is', function($m) use(&$styles) {

			$m[1] = preg_replace('#/\*.*?\*/|@[^{}]*\{.*?\}#s', '', $m[1]);

			if (preg_match_all('#([\w_.\#,\s\-\*]+)\s*{(.+?)\}#s', $m[1], $matches))
			{
				foreach($matches[1] as $key => $match)
				{
					if ($style = trim($matches[2][$key]))
					{
						$selectors = array_filter(array_map('trim', explode(',', $match)));

						foreach($selectors as $selector)
						{
							if (mb_substr($selector, -1, 1) !== '.')
							{
								$styles[] = [$selector, $style];
							}
						}
					}
				}
			}

			return $m[0];

		}, $contents);

		$html = new DOMDocument;
		$html->loadHtml($contents);

		$tagsMap = [];
		$classesMap = [];
		$idsMap = [];

		$nodes = $this->getElements($html, '*');

		foreach($nodes as $element)
		{
			$name = strtolower($element->tagName);

			if ($name === 'head' || $element->parentNode && !strcasecmp('head', $element->parentNode->nodeName))
			{
				continue;
			}

			if (!isset($tagsMap[$name]))
			{
				$tagsMap[$name] = [];
			}

			$tagsMap[$name][] = $element;

			if ($classes = $element->getAttribute('class'))
			{
				$classes = array_filter(array_map('trim', explode(' ', $classes)));

				foreach($classes as $class)
				{
					if (!isset($classesMap[$class]))
					{
						$classesMap[$class] = [];
					}

					$classesMap[$class][] = $element;
				}
			}

			if ($id = $element->getAttribute('id'))
			{
				$idsMap[$id] = $element;
			}
		}

		foreach(array_reverse($styles) as $style)
		{
			if (($pos = strpos($style[0], '.')) !== false)
			{
				$class = substr($style[0], $pos + 1);

				if (isset($classesMap[$class]))
				{
					foreach($classesMap[$class] as $element)
					{
						if (!$pos || !strcasecmp(substr($style[0], 0, $pos), $element->tagName))
						{
							$element->setAttribute('style', $style[1] . ';' . $element->getAttribute('style'));
						}
					}
				}
			}
			if (($pos = strpos($style[0], '#')) !== false)
			{
				$id = substr($style[0], $pos + 1);

				if ($element = $idsMap[$id] ?? null)
				{
					if (!$pos || !strcasecmp(substr($style[0], 0, $pos), $element->tagName))
					{
						$element->setAttribute('style', $style[1] . ';' . $element->getAttribute('style'));
					}
				}
			}
			else
			{
				$tagName = strtolower($style[0]);

				if ($tagName === '*')
				{
					foreach($tagsMap as $array)
					{
						foreach($array as $element)
						{
							$element->setAttribute('style', $style[1] . ';' . $element->getAttribute('style'));
						}
					}
				}
				else if (isset($tagsMap[$tagName]))
				{
					foreach($tagsMap[$tagName] as $element)
					{
						$element->setAttribute('style', $style[1] . ';' . $element->getAttribute('style'));
					}
				}
			}
		}

		$tagsMap['body'][0]->setAttribute('style', $tagsMap['body'][0]->getAttribute('style') . '; margin: 0;');

		return "<!DOCTYPE HTML>\n" . $html->saveHtml($html->documentElement);
	}


	/**
	 * Error handler. Can be replaced with the callable argument in the constructor.
	 *
	 * @param   string   $error    Error text.
	 * @param   string   $status   Error unique status.
	 *
	 * @return  void
	 *
	 */
	protected function raiseError(string $error, string $status): void
	{
		if ($this->errorHandler)
		{
			call_user_func($this->errorHandler, $error, $status);
		}
		else
		{
			echo $error;
			http_response_code(404);
		}

		exit;
	}
}
