import * as React from 'react';
import { useState, useRef, useEffect } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { LIST_IDS } from '../../constants/ListConstant';

// PnP JS v4 Imports
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/attachments';
import '@pnp/sp/site-users';

// People Picker Control Imports
import { PeoplePicker, PrincipalType, IPeoplePickerContext } from '@pnp/spfx-controls-react/lib/PeoplePicker';

export interface IFormProps {
  context: WebPartContext;
}

export const Form: React.FC<IFormProps> = ({ context }) => {
  // Form Field States
  const [schoolName, setSchoolName] = useState<string>('');
  const [selectedSchool, setSelectedSchool] = useState<string>('');
  const [selectedPersonId, setSelectedPersonId] = useState<number | null>(null);
  const [selectedPersonLoginName, setSelectedPersonLoginName] = useState<string>('');
  const [selectedPersonEmail, setSelectedPersonEmail] = useState<string>('');
  const [documentStatus, setDocumentStatus] = useState<string>('New');
  const [nonPecuniaryDamages, setNonPecuniaryDamages] = useState<string>('');
  const [punitiveDamages, setPunitiveDamages] = useState<string>('');
  const [comments, setComments] = useState<string>('');
  const [attachmentFiles, setAttachmentFiles] = useState<File[]>([]);

  // UI Control States
  const [showSchoolDropdown, setShowSchoolDropdown] = useState<boolean>(false);
  const [isSubmitting, setIsSubmitting] = useState<boolean>(false);
  const [showToast, setShowToast] = useState<boolean>(false);
  const [toastMessage, setToastMessage] = useState<string>('');
  const [toastType, setToastType] = useState<'success' | 'error'>('success');
  const [azureResponse, setAzureResponse] = useState<any | null>(null);
  
  const fileInputRef = useRef<HTMLInputElement>(null);
  const schoolDropdownRef = useRef<HTMLDivElement>(null);

  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient
};

  // Close dropdowns on outside click
  useEffect(() => {
    function handleClickOutside(event: MouseEvent): void {
      if (schoolDropdownRef.current && !schoolDropdownRef.current.contains(event.target as Node)) {
        setShowSchoolDropdown(false);
      }
    }
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const handleAddAttachmentClick = (): void => {
    if (fileInputRef.current) {
      fileInputRef.current.click();
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    if (e.target.files && e.target.files.length > 0) {
      const newFiles: File[] = [];
      for (let i = 0; i < e.target.files.length; i++) {
        newFiles.push(e.target.files[i]);
      }
      setAttachmentFiles(prev => [...prev, ...newFiles]);
    }
  };

  const handleRemoveAttachment = (indexToRemove: number): void => {
    setAttachmentFiles(prev => prev.filter((_, idx) => idx !== indexToRemove));
  };

  const _getPeoplePickerItems = (items: any[]): void => {
    if (items && items.length > 0) {
      setSelectedPersonLoginName(items[0].loginName || items[0].id || '');
      setSelectedPersonEmail(items[0].secondaryText || items[0].loginName || '');
    } else {
      setSelectedPersonLoginName('');
      setSelectedPersonEmail('');
    }
  };

  const handleSave = async (): Promise<void> => {
    if (!schoolName.trim()) {
      setToastType('error');
      setToastMessage('Please enter the name of the Schools.');
      setShowToast(true);
      setTimeout(() => setShowToast(false), 3000);
      return;
    }

    setIsSubmitting(true);
    setAzureResponse(null);
    try {
      // Initialize PnP JS
      const sp = spfi().using(SPFx(context));

      let userId: number | null = null;
      if (selectedPersonLoginName) {
        const userResult = await sp.web.ensureUser(selectedPersonLoginName);
        userId = userResult.Id;
      }

      // Construct item payload
      const payload: Record<string, any> = {
        Title: schoolName,
        Schools: selectedSchool,
        DocumentStatus: documentStatus,
        NonPecuniaryDamages: nonPecuniaryDamages,
        PunitiveDamages: punitiveDamages,
        SchoolsComments: comments
      };

      // Set AssignTo Person Field
      if (userId) {
        payload.AssignToId = userId;
      }

      // 1. Add item to SharePoint List
      const list = sp.web.lists.getById(LIST_IDS.Schools);
      const itemResult = await list.items.add(payload);
      const itemId = itemResult.Id;

      // 2. Upload attachments if any
      let uploadedAttachments: { name: string; url: string }[] = [];
      if (attachmentFiles.length > 0) {
        const item = list.items.getById(itemId);
        for (const file of attachmentFiles) {
          await item.attachmentFiles.add(file.name, file);
        }
        try {
          const attachmentsInfo = await item.attachmentFiles();
          uploadedAttachments = attachmentsInfo.map(att => ({
            name: att.FileName,
            url: att.ServerRelativeUrl
          }));
        } catch (attErr) {
          console.error('Error fetching uploaded attachments info:', attErr);
          // Fallback manual URL generation
          uploadedAttachments = attachmentFiles.map(file => ({
            name: file.name,
            url: `${context.pageContext.web.serverRelativeUrl}/Lists/Schools/Attachments/${itemId}/${file.name}`
          }));
        }
      }

      // 3. POST to Azure Function
      try {
        const response = await fetch('http://localhost:7071/api/mcpDemo', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({
            listId: LIST_IDS.Schools,
            itemId,
            schoolName,
            selectedSchool,
            assignToLogin: selectedPersonLoginName,
            assignToEmail: selectedPersonEmail,
            documentStatus,
            nonPecuniaryDamages,
            punitiveDamages,
            comments,
            attachments: uploadedAttachments
          })
        });

        if (response.ok) {
          const resData = await response.json();
          setAzureResponse(resData);
        } else {
          console.error('Azure Function returned status:', response.status);
        }
      } catch (azureErr) {
        console.error('Error calling Azure Function:', azureErr);
      }

      setToastType('success');
      setToastMessage('Item saved to SharePoint and posted to Azure Function!');
      setShowToast(true);
    } catch (error: any) {
      console.error('Error saving item to SharePoint list:', error);
      setToastType('error');
      setToastMessage(`Failed to save item: ${error.message || String(error)}`);
      setShowToast(true);
    } finally {
      setIsSubmitting(false);
      setTimeout(() => setShowToast(false), 5000);
    }
  };

  const handleCancel = (): void => {
    setSchoolName('');
    setSelectedSchool('');
    setSelectedPersonId(null);
    setSelectedPersonLoginName('');
    setSelectedPersonEmail('');
    setDocumentStatus('New');
    setNonPecuniaryDamages('');
    setPunitiveDamages('');
    setComments('');
    setAttachmentFiles([]);
    setAzureResponse(null);
  };

  return (
    <div className="max-w-4xl mx-auto my-6 bg-white border border-gray-200 rounded-lg shadow-sm font-sans text-gray-800 relative">
      
      {/* Toast Notification */}
      {showToast && (
        <div className={`fixed top-10 right-4 ${toastType === 'success' ? 'bg-emerald-600' : 'bg-rose-600'} text-white px-4 py-3 rounded shadow-lg z-50 transition-opacity duration-300 flex items-center space-x-2`}>
          {toastType === 'success' ? (
            <svg className="w-5 h-5 fill-current" viewBox="0 0 20 20">
              <path d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" />
            </svg>
          ) : (
            <svg className="w-5 h-5 fill-current" viewBox="0 0 20 20">
              <path fillRule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7 4a1 1 0 11-2 0 1 1 0 012 0zm-1-9a1 1 0 00-1 1v4a1 1 0 102 0V6a1 1 0 00-1-1z" clipRule="evenodd" />
            </svg>
          )}
          <span className="text-sm font-medium">{toastMessage}</span>
        </div>
      )}

      {/* Top Navigation / Action Bar */}
      <div className="flex items-center justify-between px-6 py-3 border-b border-gray-200 bg-gray-50/50">
        <div className="flex items-center space-x-4">
          <span className="text-lg font-semibold text-gray-900">New item</span>
          {isSubmitting && (
            <span className="text-xs text-indigo-600 font-medium animate-pulse flex items-center">
              <svg className="animate-spin h-4 w-4 mr-2 text-indigo-600" fill="none" viewBox="0 0 24 24">
                <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" />
              </svg>
              Saving to SharePoint...
            </span>
          )}
        </div>
        <div className="flex items-center space-x-6">
          <button onClick={handleCancel} disabled={isSubmitting} className="text-gray-400 hover:text-gray-600 transition-colors disabled:opacity-50">
            <svg className="w-5 h-5" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12" />
            </svg>
          </button>
        </div>
      </div>

      {/* Form Content Area */}
      <div className="p-6 space-y-6">

        {/* Azure Function Response */}
        {azureResponse && (
          <div className="p-4 bg-emerald-50 border border-emerald-200 rounded-lg text-emerald-800 space-y-2">
            <h3 className="font-semibold text-sm flex items-center">
              <svg className="w-5 h-5 mr-2 text-emerald-600 fill-current" viewBox="0 0 20 20">
                <path d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" />
              </svg>
              Response from Azure Function:
            </h3>
            <p className="text-xs">{azureResponse.message}</p>
            <div className="text-xs bg-white/80 p-3 rounded border border-emerald-100 font-mono text-gray-700 max-h-48 overflow-y-auto">
              <pre>{JSON.stringify(azureResponse.receivedData, null, 2)}</pre>
            </div>
          </div>
        )}

        {/* Field 1: Please enter the name of the Schools */}
        <div className="space-y-1.5">
          <label className="flex items-center text-xs font-semibold text-gray-700 tracking-wide uppercase">
            <span className="inline-flex items-center justify-center w-5 h-5 mr-2 rounded border border-gray-400 text-[10px] font-bold bg-white text-gray-500">
              T
            </span>
            Please enter the name of the Schools
          </label>
          <input
            type="text"
            placeholder="Enter value here"
            value={schoolName}
            disabled={isSubmitting}
            onChange={(e) => setSchoolName(e.target.value)}
            className="w-full px-3 py-2 text-sm border border-gray-300 rounded focus:outline-none focus:border-indigo-600 focus:ring-1 focus:ring-indigo-600 bg-gray-50/50 hover:bg-gray-50 transition-colors disabled:opacity-60"
          />
        </div>

        {/* Field 2: Schools Dropdown */}
        <div className="space-y-1.5 relative" ref={schoolDropdownRef}>
          <label className="flex items-center text-xs font-semibold text-gray-700 tracking-wide uppercase">
            <span className="inline-flex items-center justify-center w-5 h-5 mr-2 text-gray-500">
              <svg className="w-4 h-4" fill="none" stroke="currentColor" strokeWidth="2.5" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
              </svg>
            </span>
            Schools
          </label>
          <div
            onClick={() => !isSubmitting && setShowSchoolDropdown(!showSchoolDropdown)}
            className="w-full px-3 py-2 text-sm border border-gray-300 rounded cursor-pointer flex justify-between items-center bg-gray-50/50 hover:bg-gray-50 transition-colors disabled:opacity-60"
          >
            <span className={selectedSchool ? 'text-gray-900' : 'text-gray-400'}>
              {selectedSchool || '—'}
            </span>
            <svg className="w-4 h-4 text-gray-500" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" d="M19 9l-7 7-7-7" />
            </svg>
          </div>
          {showSchoolDropdown && (
            <div className="absolute left-0 z-10 w-full mt-1 bg-white border border-gray-200 rounded shadow-lg max-h-48 overflow-y-auto">
              <div 
                onClick={() => { setSelectedSchool(''); setShowSchoolDropdown(false); }}
                className="px-3 py-2 text-sm text-gray-500 hover:bg-indigo-50 hover:text-indigo-900 cursor-pointer"
              >
                —
              </div>
              {['Greenwood High', 'Oakridge International', 'Silver Oaks School', 'Orchids International School', 'Delhi Public School'].map((school) => (
                <div
                  key={school}
                  onClick={() => {
                    setSelectedSchool(school);
                    setShowSchoolDropdown(false);
                  }}
                  className="px-3 py-2 text-sm hover:bg-indigo-50 hover:text-indigo-900 cursor-pointer text-gray-800"
                >
                  {school}
                </div>
              ))}
            </div>
          )}
        </div>

        {/* Field 3: AssignTo People Picker */}
        <div className="space-y-1.5 relative">
          <label className="flex items-center text-xs font-semibold text-gray-700 tracking-wide uppercase mb-1">
            <span className="inline-flex items-center justify-center w-5 h-5 mr-2 text-gray-500">
              <svg className="w-4.5 h-4.5" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" d="M16 7a4 4 0 11-8 0 4 4 0 018 0zM12 14a7 7 0 00-7 7h14a7 7 0 00-7-7z" />
              </svg>
            </span>
            AssignTo
          </label>
          <PeoplePicker
            context={peoplePickerContext}
            personSelectionLimit={1}
            showtooltip={true}
            required={false}
            disabled={isSubmitting}
            onChange={_getPeoplePickerItems}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
            placeholder="Enter a name or email address"
          />
        </div>

        {/* Field 4: Document Status */}
        <div className="space-y-1.5">
          <label className="flex items-center text-xs font-semibold text-gray-700 tracking-wide uppercase">
            <span className="inline-flex items-center justify-center w-5 h-5 mr-2 text-gray-500">
              <svg className="w-4 h-4" fill="none" stroke="currentColor" strokeWidth="2.5" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
              </svg>
            </span>
            Document Status
          </label>
          <div className="w-full px-3 py-2 border border-gray-300 rounded bg-gray-50/50 flex items-center space-x-2">
            <select
              value={documentStatus}
              disabled={isSubmitting}
              onChange={(e) => setDocumentStatus(e.target.value)}
              className="text-sm bg-transparent border-none text-gray-800 focus:ring-0 focus:outline-none cursor-pointer w-full disabled:opacity-60"
            >
              <option value="New">New</option>
              <option value="In Progress">In Progress</option>
              <option value="Approved">Approved</option>
              <option value="Archived">Archived</option>
            </select>
          </div>
        </div>

        {/* Field 5: Please enter the quantum for Non-Pecuniary and Special Damages */}
        <div className="space-y-1.5">
          <label className="flex items-center text-xs font-semibold text-gray-700 tracking-wide uppercase">
            <span className="inline-flex items-center justify-center w-5 h-5 mr-2 rounded border border-gray-400 text-[10px] font-bold bg-white text-gray-500">
              T
            </span>
            Please enter the quantum for Non-Pecuniary and Special Damages
          </label>
          <input
            type="text"
            placeholder="Enter value here"
            value={nonPecuniaryDamages}
            disabled={isSubmitting}
            onChange={(e) => setNonPecuniaryDamages(e.target.value)}
            className="w-full px-3 py-2 text-sm border border-gray-300 rounded focus:outline-none focus:border-indigo-600 focus:ring-1 focus:ring-indigo-600 bg-gray-50/50 hover:bg-gray-50 transition-colors disabled:opacity-60"
          />
        </div>

        {/* Field 6: Please enter the quantum for Punitive and Aggravated Damages */}
        <div className="space-y-1.5">
          <label className="flex items-center text-xs font-semibold text-gray-700 tracking-wide uppercase">
            <span className="inline-flex items-center justify-center w-5 h-5 mr-2 rounded border border-gray-400 text-[10px] font-bold bg-white text-gray-500">
              T
            </span>
            Please enter the quantum for Punitive and Aggravated Damages
          </label>
          <input
            type="text"
            placeholder="Enter value here"
            value={punitiveDamages}
            disabled={isSubmitting}
            onChange={(e) => setPunitiveDamages(e.target.value)}
            className="w-full px-3 py-2 text-sm border border-gray-300 rounded focus:outline-none focus:border-indigo-600 focus:ring-1 focus:ring-indigo-600 bg-gray-50/50 hover:bg-gray-50 transition-colors disabled:opacity-60"
          />
        </div>

        {/* Field 7: Schools Comments */}
        <div className="space-y-1.5">
          <label className="flex items-center text-xs font-semibold text-gray-700 tracking-wide uppercase">
            <span className="inline-flex items-center justify-center w-5 h-5 mr-2 text-gray-500">
              <svg className="w-4 h-4" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" d="M4 6h16M4 12h16M4 18h12" />
              </svg>
            </span>
            Schools Comments
          </label>
          <textarea
            placeholder="Enter value here"
            value={comments}
            rows={4}
            disabled={isSubmitting}
            onChange={(e) => setComments(e.target.value)}
            className="w-full px-3 py-2 text-sm border border-gray-300 rounded focus:outline-none focus:border-indigo-600 focus:ring-1 focus:ring-indigo-600 bg-gray-50/50 hover:bg-gray-50 transition-colors resize-y disabled:opacity-60"
          />
        </div>

        {/* Field 8: Attachments */}
        <div className="space-y-1.5">
          <label className="flex items-center text-xs font-semibold text-gray-700 tracking-wide uppercase">
            <span className="inline-flex items-center justify-center w-5 h-5 mr-2 text-gray-500">
              <svg className="w-4.5 h-4.5" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" d="M15.172 7l-6.586 6.586a2 2 0 102.828 2.828l6.414-6.586a4 4 0 00-5.656-5.656l-6.415 6.585a6 6 0 108.486 8.486L20.5 13" />
              </svg>
            </span>
            Attachments
          </label>
          <input
            type="file"
            ref={fileInputRef}
            disabled={isSubmitting}
            onChange={handleFileChange}
            multiple
            className="hidden"
          />
          <div className="w-full px-3 py-3 border border-gray-300 rounded bg-gray-50/50 flex flex-col space-y-2">
            <div>
              <button
                type="button"
                onClick={handleAddAttachmentClick}
                disabled={isSubmitting}
                className="text-indigo-700 hover:text-indigo-900 text-sm font-medium inline-flex items-center transition-colors disabled:opacity-50"
              >
                Add attachments
              </button>
            </div>
            
            {attachmentFiles.length > 0 && (
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 pt-2">
                {attachmentFiles.map((file, idx) => (
                  <div key={idx} className="flex items-center justify-between p-2 bg-white border border-gray-200 rounded text-xs">
                    <span className="truncate text-gray-700 font-medium max-w-[200px]" title={file.name}>
                      {file.name} ({(file.size / 1024).toFixed(1)} KB)
                    </span>
                    <button
                      type="button"
                      onClick={() => handleRemoveAttachment(idx)}
                      disabled={isSubmitting}
                      className="text-red-500 hover:text-red-700 font-semibold focus:outline-none disabled:opacity-50"
                    >
                      Delete
                    </button>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>

      </div>

      {/* Footer Area with Save and Cancel */}
      <div className="flex justify-end space-x-3 px-6 py-4 border-t border-gray-200 bg-gray-50/50">
        <button
          onClick={handleSave}
          disabled={isSubmitting}
          className="px-6 py-1.5 bg-indigo-700 hover:bg-indigo-800 text-white font-medium text-sm rounded shadow-sm hover:shadow transition-all disabled:opacity-50 flex items-center"
        >
          {isSubmitting && (
            <svg className="animate-spin -ml-1 mr-2 h-4 w-4 text-white" fill="none" viewBox="0 0 24 24">
              <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
              <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" />
            </svg>
          )}
          Save
        </button>
        <button
          onClick={handleCancel}
          disabled={isSubmitting}
          className="px-6 py-1.5 border border-gray-300 hover:bg-gray-100 text-gray-700 font-medium text-sm rounded transition-colors disabled:opacity-50"
        >
          Cancel
        </button>
      </div>
    </div>
  );
};
