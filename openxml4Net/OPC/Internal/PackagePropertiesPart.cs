﻿using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using NPOI.OpenXml4Net.Exceptions;
using NPOI.OpenXml4Net.OPC;
using NPOI.SS.Util;
using System.Globalization;

namespace NPOI.OpenXml4Net.OPC.Internal
{
/**
 * Represents the core properties part of a package.
 * 
 * @author Julien Chable
 * @version 1.0
 */
public class PackagePropertiesPart:PackagePart,PackageProperties 
{
    static string NAMESPACE_DC = "http://purl.org/dc/elements/1.1/";

	public static string NAMESPACE_DC_URI = "http://purl.org/dc/elements/1.1/";

	public static string NAMESPACE_CP_URI = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";

	public static string NAMESPACE_DCTERMS_URI = "http://purl.org/dc/terms/";

	public static string NAMESPACE_XSI_URI = "http://www.w3.org/2001/XMLSchema-instance";

	/**
	 * Constructor.
	 * 
	 * @param pack
	 *            Container package.
	 * @param partName
	 *            Name of this part.
	 * @throws InvalidFormatException
	 *             Throws if the content is invalid.
	 */
	public PackagePropertiesPart(OPCPackage pack, PackagePartName partName)
        :base(pack, partName, ContentTypes.CORE_PROPERTIES_PART)
	{
		
	}

	/**
	 * A categorization of the content of this package.
	 * 
	 * [Example: Example values for this property might include: Resume, Letter,
	 * Financial Forecast, Proposal, Technical Presentation, and so on. This
	 * value might be used by an application's user interface to facilitate
	 * navigation of a large Set of documents. end example]
	 */
	protected string category = null;

	/**
	 * The status of the content.
	 * 
	 * [Example: Values might include "Draft", "Reviewed", and "Final". end
	 * example]
	 */
	protected string contentStatus = null;

	/**
	 * The type of content represented, generally defined by a specific use and
	 * intended audience.
	 * 
	 * [Example: Values might include "Whitepaper", "Security Bulletin", and
	 * "Exam". end example] [Note: This property is distinct from MIME content
	 * types as defined in RFC 2616. end note]
	 */
	protected string contentType = null;

	/**
	 * Date of creation of the resource.
	 */
	protected Nullable<DateTime> created = new Nullable<DateTime>();

	/**
	 * An entity primarily responsible for making the content of the resource.
	 */
	protected string creator = null;

	/**
	 * An explanation of the content of the resource.
	 * 
	 * [Example: Values might include an abstract, table of contents, reference
	 * to a graphical representation of content, and a free-text account of the
	 * content. end example]
	 */
	protected string description = null;

	/**
	 * An unambiguous reference to the resource within a given context.
	 */
	protected string identifier = null;

	/**
	 * A delimited Set of keywords to support searching and indexing. This is
	 * typically a list of terms that are not available elsewhere in the
	 * properties.
	 */
	protected string keywords = null;

	/**
	 * The language of the intellectual content of the resource.
	 * 
	 * [Note: IETF RFC 3066 provides guidance on encoding to represent
	 * languages. end note]
	 */
	protected string language = null;

	/**
	 * The user who performed the last modification. The identification is
	 * environment-specific.
	 * 
	 * [Example: A name, email address, or employee ID. end example] It is
	 * recommended that this value be as concise as possible.
	 */
	protected string lastModifiedBy = null;

	/**
	 * The date and time of the last printing.
	 */
	protected Nullable<DateTime> lastPrinted = new Nullable<DateTime>();

	/**
	 * Date on which the resource was changed.
	 */
	protected Nullable<DateTime> modified = new Nullable<DateTime>();

	/**
	 * The revision number.
	 * 
	 * [Example: This value might indicate the number of saves or revisions,
	 * provided the application updates it after each revision. end example]
	 */
	protected string revision = null;

	/**
	 * The topic of the content of the resource.
	 */
	protected string subject = null;

	/**
	 * The name given to the resource.
	 */
	protected string title = null;

	/**
	 * The version number. This value is Set by the user or by the application.
	 */
	protected string version = null;

	/*
	 * Getters and Setters
	 */

	/**
	 * Get the category property.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#getCategoryProperty()
	 */
	public string GetCategoryProperty() {
		return category;
	}

	/**
	 * Get content status.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#getContentStatusProperty()
	 */
	public string GetContentStatusProperty() {
		return contentStatus;
	}

	/**
	 * Get content type.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#getContentTypeProperty()
	 */
	public string GetContentTypeProperty() {
		return contentType;
	}

	/**
	 * Get created date.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#getCreatedProperty()
	 */
	public Nullable<DateTime> GetCreatedProperty() {
		return created;
	}

	/**
	 * Get created date formated into a String.
	 * 
	 * @return A string representation of the created date.
	 */
	public string GetCreatedPropertyString() {
		return GetDateValue(created);
	}

	/**
	 * Get creator.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#getCreatorProperty()
	 */
	public string GetCreatorProperty() {
		return creator;
	}

	/**
	 * Get description.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#getDescriptionProperty()
	 */
	public string GetDescriptionProperty() {
		return description;
	}

	/**
	 * Get identifier.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#getIdentifierProperty()
	 */
	public string GetIdentifierProperty() {
		return identifier;
	}

	/**
	 * Get keywords.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#getKeywordsProperty()
	 */
	public string GetKeywordsProperty() {
		return keywords;
	}

	/**
	 * Get the language.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#getLanguageProperty()
	 */
	public string GetLanguageProperty() {
		return language;
	}

	/**
	 * Get the author of last modifications.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#getLastModifiedByProperty()
	 */
	public string GetLastModifiedByProperty() {
		return lastModifiedBy;
	}

	/**
	 * Get last printed date.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#getLastPrintedProperty()
	 */
	public Nullable<DateTime> GetLastPrintedProperty() {
		return lastPrinted;
	}

	/**
	 * Get last printed date formated into a String.
	 * 
	 * @return A string representation of the last printed date.
	 */
	public string GetLastPrintedPropertyString() {
        return GetDateValue(lastPrinted);
	}

	/**
	 * Get modified date.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#getModifiedProperty()
	 */
	public Nullable<DateTime> GetModifiedProperty() {
		return modified;
	}

	/**
	 * Get modified date formated into a String.
	 * 
	 * @return A string representation of the modified date.
	 */
	public string GetModifiedPropertyString() {
		if (modified.Value==null)
			return GetDateValue(new Nullable<DateTime>(new DateTime()));
		else
			return GetDateValue(modified);
	}

	/**
	 * Get revision.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#getRevisionProperty()
	 */
	public string GetRevisionProperty() {
		return revision;
	}

	/**
	 * Get subject.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#getSubjectProperty()
	 */
	public string GetSubjectProperty() {
		return subject;
	}

	/**
	 * Get title.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#getTitleProperty()
	 */
	public string GetTitleProperty() {
		return title;
	}

	/**
	 * Get version.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#getVersionProperty()
	 */
	public string GetVersionProperty() {
		return version;
	}

	/**
	 * Set the category.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#setCategoryProperty(java.lang.String)
	 */
	public void SetCategoryProperty(string category) {
		this.category = SetStringValue(category);
	}

	/**
	 * Set the content status.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#setContentStatusProperty(java.lang.String)
	 */
	public void SetContentStatusProperty(string contentStatus) {
		this.contentStatus = SetStringValue(contentStatus);
	}

	/**
	 * Set the content type.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#setContentTypeProperty(java.lang.String)
	 */
	public void SetContentTypeProperty(string contentType) {
		this.contentType = SetStringValue(contentType);
	}

	/**
	 * Set the created date.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#setCreatedProperty(org.apache.poi.OpenXml4Net.util.Nullable)
	 */
	public void SetCreatedProperty(string created) {
		try {
			this.created = SetDateValue(created);
		} catch (InvalidFormatException e) {
			new ArgumentException("created  : "
					+ e.Message);
		}
	}

	/**
	 * Set the created date.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#setCreatedProperty(org.apache.poi.OpenXml4Net.util.Nullable)
	 */
	public void SetCreatedProperty(Nullable<DateTime> created) {
		if (created!=null)
			this.created = created;
	}

	/**
	 * Set the creator.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#setCreatorProperty(java.lang.String)
	 */
	public void SetCreatorProperty(string creator) {
		this.creator = SetStringValue(creator);
	}

	/**
	 * Set the description.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#setDescriptionProperty(java.lang.String)
	 */
	public void SetDescriptionProperty(string description) {
		this.description = SetStringValue(description);
	}

	/**
	 * Set identifier.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#setIdentifierProperty(java.lang.String)
	 */
	public void SetIdentifierProperty(string identifier) {
		this.identifier = SetStringValue(identifier);
	}

	/**
	 * Set keywords.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#setKeywordsProperty(java.lang.String)
	 */
	public void SetKeywordsProperty(string keywords) {
		this.keywords = SetStringValue(keywords);
	}

	/**
	 * Set language.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#setLanguageProperty(java.lang.String)
	 */
	public void SetLanguageProperty(string language) {
		this.language = SetStringValue(language);
	}

	/**
	 * Set last modifications author.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#setLastModifiedByProperty(java.lang.String)
	 */
	public void SetLastModifiedByProperty(string lastModifiedBy) {
		this.lastModifiedBy = SetStringValue(lastModifiedBy);
	}

	/**
	 * Set last printed date.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#setLastPrintedProperty(org.apache.poi.OpenXml4Net.util.Nullable)
	 */
	public void SetLastPrintedProperty(string lastPrinted) {
		try {
			this.lastPrinted = SetDateValue(lastPrinted);
		} catch (InvalidFormatException e) {
			new ArgumentException("lastPrinted  : "
					+ e.Message);
		}
	}

	/**
	 * Set last printed date.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#setLastPrintedProperty(org.apache.poi.OpenXml4Net.util.Nullable)
	 */
	public void SetLastPrintedProperty(Nullable<DateTime> lastPrinted) {
		if (lastPrinted!=null)
			this.lastPrinted = lastPrinted;
	}

	/**
	 * Set last modification date.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#setModifiedProperty(org.apache.poi.OpenXml4Net.util.Nullable)
	 */
	public void SetModifiedProperty(string modified) {
		try {
			this.modified = SetDateValue(modified);
		} catch (InvalidFormatException e) {
			new ArgumentException("modified  : "
					+ e.Message);
		}
	}

	/**
	 * Set last modification date.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#setModifiedProperty(org.apache.poi.OpenXml4Net.util.Nullable)
	 */
	public void SetModifiedProperty(Nullable<DateTime> modified) {
		if (modified.HasValue)
			this.modified = modified;
	}

	/**
	 * Set revision.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#setRevisionProperty(java.lang.String)
	 */
	public void SetRevisionProperty(string revision) {
		this.revision = SetStringValue(revision);
	}

	/**
	 * Set subject.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#setSubjectProperty(java.lang.String)
	 */
	public void SetSubjectProperty(string subject) {
		this.subject = SetStringValue(subject);
	}

	/**
	 * Set title.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#setTitleProperty(java.lang.String)
	 */
	public void SetTitleProperty(string title) {
		this.title = SetStringValue(title);
	}

	/**
	 * Set version.
	 * 
	 * @see org.apache.poi.OpenXml4Net.opc.PackageProperties#setVersionProperty(java.lang.String)
	 */
	public void SetVersionProperty(string version) {
		this.version = SetStringValue(version);
	}

	/**
	 * Convert a strig value into a String
	 */
	private string SetStringValue(string s) {
		if (s == null || s.Equals(""))
			return null;
		else
			return s;
	}

	/**
	 * Convert a string value represented a date into a Nullable<DateTime>.
	 * 
	 * @throws InvalidFormatException
	 *             Throws if the date format isnot valid.
	 */
	private Nullable<DateTime> SetDateValue(string s){
		if (s == null || s.Equals(""))
			return new Nullable<DateTime>();
		else {
			SimpleDateFormat df = new SimpleDateFormat(
					"yyyy-MM-dd'T'HH:mm:ss'Z'");
			DateTime d = (DateTime)df.ParseObject(s, 0);
			if (d == null)
				throw new InvalidFormatException("Date not well formated");
			return new Nullable<DateTime>(d);
		}
	}

	/**
	 * Convert a Nullable<DateTime> into a String.
	 * 
	 * @param d
	 *            The Date to convert.
	 * @return The formated date or null.
	 * @see java.util.SimpleDateFormat
	 */
	private string GetDateValue(Nullable<DateTime> d) {
		if (d == null || d.Equals(""))
			return "";
		else {
			SimpleDateFormat df = new SimpleDateFormat(
					"yyyy-MM-dd'T'HH:mm:ss'Z'");
            return df.Format(d.Value, CultureInfo.CurrentCulture);
		}
	}

	
	protected override Stream GetInputStreamImpl() {
		throw new InvalidOperationException("Operation not authorized");
	}

    protected override Stream GetOutputStreamImpl()
    {
        throw new InvalidOperationException("Operation not authorized");
    }

	
	public override bool Save(Stream zos) {
		throw new InvalidOperationException("Operation not authorized");
	}

	
	public override bool Load(Stream ios) {
		throw new InvalidOperationException("Operation not authorized");
	}

	
	public override void Close() {
		// Do nothing
	}

	
	public override void Flush() {
		// Do nothing
	}
}

}
