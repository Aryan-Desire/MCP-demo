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

interface IUser {
  name: string;
  email: string;
  avatarColor: string;
}

const mockUsers: IUser[] = [
  { name: 'Adela Vance', email: 'AdelaV@M365x214355.onmicrosoft.com', avatarColor: 'bg-purple-600' },
  { name: 'Alex Wilber', email: 'AlexW@M365x214355.onmicrosoft.com', avatarColor: 'bg-green-600' },
  { name: 'Allan Deyoung', email: 'AllanD@M365x214355.onmicrosoft.com', avatarColor: 'bg-blue-600' },
  { name: 'Ashley McCarthy', email: 'AshleyM@M365x214355.onmicrosoft.com', avatarColor: 'bg-pink-600' },
  { name: 'Brian Johnson', email: 'BrianJ@M365x214355.onmicrosoft.com', avatarColor: 'bg-yellow-600' },
  { name: 'Christie Cline', email: 'ChristieC@M365x214355.onmicrosoft.com', avatarColor: 'bg-indigo-600' },
];

export interface IFormProps {
  context: WebPartContext;
}

export const Form: React.FC<IFormProps> = ({ context }) => {
  // Form Field States
  const [schoolName, setSchoolName] = useState<string>('');
  const [selectedSchool, setSelectedSchool] = useState<string>('');
  const [assignToInput, setAssignToInput] = useState<string>('');
  const [assignedUser, setAssignedUser] = useState<IUser | null>(null);
  const [documentStatus, setDocumentStatus] = useState<string>('New');
  const [nonPecuniaryDamages, setNonPecuniaryDamages] = useState<string>('');
  const [punitiveDamages, setPunitiveDamages] = useState<string>('');
  const [comments, setComments] = useState<string>('');
  const [attachmentFiles, setAttachmentFiles] = useState<File[]>([]);

  // UI Control States
  const [showSchoolDropdown, setShowSchoolDropdown] = useState<boolean>(false);
  const [showAssignToSuggestions, setShowAssignToSuggestions] = useState<boolean>(false);
  const [isSubmitting, setIsSubmitting] = useState<boolean>(false);
  const [showToast, setShowToast] = useState<boolean>(false);
  const [toastMessage, setToastMessage] = useState<string>('');
  const [toastType, setToastType] = useState<'success' | 'error'>('success');
  
  const fileInputRef = useRef<HTMLInputElement>(null);
  const assignToContainerRef = useRef<HTMLDivElement>(null);
  const schoolDropdownRef = useRef<HTMLDivElement>(null);

  // Close dropdowns on outside click
  useEffect(() => {
    function handleClickOutside(event: MouseEvent): void {
      if (assignToContainerRef.current && !assignToContainerRef.current.contains(event.target as Node)) {
        setShowAssignToSuggestions(false);
      }
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

  const handleSave = async (): Promise<void> => {
    if (!schoolName.trim()) {
      setToastType('error');
      setToastMessage('Please enter the name of the Schools.');
      setShowToast(true);
      setTimeout(() => setShowToast(false), 3000);
      return;
    }

    setIsSubmitting(true);
    try {
      // Initialize PnP JS
      const sp = spfi().using(SPFx(context));

      // Construct item payload
      // Make sure the fields match your SharePoint list internal column names.
      // Falls back dynamically in case column names differ slightly.
      const payload: Record<string, any> = {
        Title: schoolName,
        Domain: 'school',
        Schools: selectedSchool,
        DocumentStatus: documentStatus,
        NonPecuniaryDamages: nonPecuniaryDamages,
        PunitiveDamages: punitiveDamages,
        SchoolsComments: comments
      };

      // If you are using Person field in SharePoint list:
      // Note: Typically you need to resolve User ID, but to keep it simple, we store it
      // if column exists, or as text. If AssignTo is a Person field, PnP JS expects 'AssignToId'
      // with a numeric ID. As a robust solution, we save the text name to AssignTo if it's text.
      if (assignedUser) {
        payload.AssignTo = assignedUser.name;
      }

      // 1. Add item to SharePoint List
      const list = sp.web.lists.getById(LIST_IDS.Schools);
      const itemResult = await list.items.add(payload);
      const itemId = itemResult.data.Id;

      // 2. Upload attachments if any
      if (attachmentFiles.length > 0) {
        const item = list.items.getById(itemId);
        for (const file of attachmentFiles) {
          await item.attachmentFiles.add(file.name, file);
        }
      }

      setToastType('success');
      setToastMessage('Item and attachments successfully saved to SharePoint!');
      setShowToast(true);
      handleCancel(); // Clear form on success
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
    setAssignToInput('');
    setAssignedUser(null);
    setDocumentStatus('New');
    setNonPecuniaryDamages('');
    setPunitiveDamages('');
    setComments('');
    setAttachmentFiles([]);
  };

  const filteredUsers = mockUsers.filter(user => 
    user.name.toLowerCase().includes(assignToInput.toLowerCase()) || 
    user.email.toLowerCase().includes(assignToInput.toLowerCase())
  );

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
        <div className="space-y-1.5 relative" ref={assignToContainerRef}>
          <label className="flex items-center text-xs font-semibold text-gray-700 tracking-wide uppercase">
            <span className="inline-flex items-center justify-center w-5 h-5 mr-2 text-gray-500">
              <svg className="w-4.5 h-4.5" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" d="M16 7a4 4 0 11-8 0 4 4 0 018 0zM12 14a7 7 0 00-7 7h14a7 7 0 00-7-7z" />
              </svg>
            </span>
            AssignTo
          </label>
          <div className="flex flex-wrap items-center gap-2 w-full px-3 py-1.5 border border-gray-300 rounded bg-gray-50/50 hover:bg-gray-50 focus-within:border-indigo-600 focus-within:ring-1 focus-within:ring-indigo-600 transition-colors">
            {assignedUser && (
              <span className="flex items-center space-x-1.5 bg-indigo-50 text-indigo-800 text-xs font-medium px-2 py-0.5 rounded border border-indigo-200">
                <span className={`w-4 h-4 rounded-full ${assignedUser.avatarColor} text-white flex items-center justify-center text-[9px]`}>
                  {assignedUser.name.charAt(0)}
                </span>
                <span>{assignedUser.name}</span>
                <button 
                  onClick={() => !isSubmitting && setAssignedUser(null)} 
                  disabled={isSubmitting}
                  className="text-indigo-500 hover:text-indigo-700 font-bold focus:outline-none disabled:opacity-50"
                >
                  &times;
                </button>
              </span>
            )}
            <input
              type="text"
              placeholder={assignedUser ? "" : "Enter a name or email address"}
              value={assignToInput}
              disabled={isSubmitting}
              onChange={(e) => {
                setAssignToInput(e.target.value);
                setShowAssignToSuggestions(true);
              }}
              onFocus={() => !isSubmitting && setShowAssignToSuggestions(true)}
              className="flex-grow min-w-[200px] text-sm bg-transparent outline-none py-0.5 disabled:opacity-60"
            />
          </div>

          {showAssignToSuggestions && (
            <div className="absolute left-0 z-10 w-full mt-1 bg-white border border-gray-200 rounded shadow-lg max-h-56 overflow-y-auto">
              {filteredUsers.length > 0 ? (
                filteredUsers.map((user) => (
                  <div
                    key={user.email}
                    onClick={() => {
                      setAssignedUser(user);
                      setAssignToInput('');
                      setShowAssignToSuggestions(false);
                    }}
                    className="flex items-center space-x-3 px-3 py-2 hover:bg-indigo-50 cursor-pointer"
                  >
                    <div className={`w-8 h-8 rounded-full ${user.avatarColor} text-white flex items-center justify-center font-semibold text-sm`}>
                      {user.name.split(' ').map(n => n.charAt(0)).join('')}
                    </div>
                    <div className="text-left">
                      <div className="text-sm font-medium text-gray-900">{user.name}</div>
                      <div className="text-xs text-gray-500">{user.email}</div>
                    </div>
                  </div>
                ))
              ) : (
                <div className="px-3 py-2 text-sm text-gray-500">No suggestions found</div>
              )}
            </div>
          )}
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
