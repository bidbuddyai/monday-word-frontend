// frontend/src/app/page.tsx
'use client';
import React, { useState, useEffect } from 'react';
import mondaySdk from 'monday-sdk-js';
import { Button } from "@/components/ui/button";
import { Card, CardHeader, CardTitle, CardContent } from "@/components/ui/card";
import { Select, SelectTrigger, SelectContent, SelectItem, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Toaster, toast } from 'sonner'; // Import Toaster and toast from sonner
import styles from './page.module.css';

const monday = mondaySdk();

interface Template {
  id: string;
  name: string;
  path: string;
  lastModified: string;
}

interface Placeholder {
  id: string;
  name: string;
  description: string;
  suggestedColumnId?: string | null;
  suggestedColumnTitle?: string | null;
}

interface Column {
  id: string;
  title: string;
  type: string;
}

export default function Home() {
  const [context, setContext] = useState<any>(null);
  const [templates, setTemplates] = useState<Template[]>([]);
  const [selectedTemplate, setSelectedTemplate] = useState<string | null>(null);
  const [placeholders, setPlaceholders] = useState<Placeholder[]>([]);
  const [columns, setColumns] = useState<Column[]>([]);
  const [placeholderMappings, setPlaceholderMappings] = useState<{ [key: string]: string }>({});
  const [loading, setLoading] = useState<boolean>(false);
  const [loadingTemplates, setLoadingTemplates] = useState<boolean>(false);
  const [generating, setGenerating] = useState<boolean>(false);
  const [error, setError] = useState<string>('');

  useEffect(() => {
    monday.listen('context', (res) => {
      setContext(res.data);
    });
  }, []);

    // --- Microsoft Login ---
    const handleMicrosoftLogin = async () => {
        setLoading(true); // General loading state
        setError('');
        try {
            const backendUrl = process.env.NEXT_PUBLIC_BACKEND_URL;

            // Ensure backend URL is defined.
            if (!backendUrl) {
                throw new Error("NEXT_PUBLIC_BACKEND_URL is not defined. Check your .env file.");
            }

            const response = await fetch(`${backendUrl}`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ action: 'get_ms_auth_url' })
            });

            if (!response.ok) {
                const errorData = await response.json(); // Attempt to get error details
                throw new Error(`Backend error: ${response.status} - ${errorData.error || "Unknown error"}`);
            }

            const data = await response.json();
            if (data.success) {
                window.location.href = data.authUrl;
            } else {
              const errorMessage = data.error || 'Failed to get Microsoft auth URL';
              setError(errorMessage);
              showErrorToast(errorMessage);
            }
        } catch (e:any) {
            console.error(e); //Always log
            setError(e.message || 'Failed to initiate Microsoft login.');
            showErrorToast(e.message || 'Failed to initiate Microsoft login.');
        } finally {
            setLoading(false); // Reset general loading
        }
    };

    // --- Load Templates ---
  const loadTemplates = async () => {
    setLoadingTemplates(true);
    setError('');
    try {
      const backendUrl = process.env.NEXT_PUBLIC_BACKEND_URL;
      if (!backendUrl) {
        throw new Error("NEXT_PUBLIC_BACKEND_URL is not defined.");
      }

      const response = await fetch(`${backendUrl}`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${context.shortLivedToken}` // Use shortLivedToken
        },
        body: JSON.stringify({ action: 'get_templates', token: context.shortLivedToken }),
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(`Backend error: ${response.status} - ${errorData.error || "Unknown error"}`);
      }

      const data = await response.json();
      if (data.success) {
          if (data.accessToken) {
              localStorage.setItem('msAccessToken', data.accessToken);
              localStorage.setItem('msRefreshToken', data.refreshToken);
              localStorage.setItem('msTokenExpiry', String(Date.now() + (data.expiresIn * 1000)));
          }
        setTemplates(data.templates);
      } else {
        setError(data.error || 'Failed to load templates');
        showErrorToast(data.error || 'Failed to load templates');
      }
    } catch (e:any) {
      console.error(e);
      setError(e.message || 'Failed to load templates.');
      showErrorToast(e.message || 'Failed to load templates.');
    } finally {
      setLoadingTemplates(false);
    }
  };

  // --- Analyze Selected Template ---
  const analyzeSelectedTemplate = async () => {
    if (!selectedTemplate) return;

    setLoading(true);
    setError('');
    try {
      const backendUrl = process.env.NEXT_PUBLIC_BACKEND_URL;
        if (!backendUrl) {
            throw new Error("NEXT_PUBLIC_BACKEND_URL is not defined.");
        }

      const accessToken = localStorage.getItem('msAccessToken');
      const refreshToken = localStorage.getItem('msRefreshToken');
      const tokenExpiry = localStorage.getItem('msTokenExpiry');

        if (!accessToken || !refreshToken || !tokenExpiry)
        {
            setError('Microsoft tokens are not available. Please log in.');
            showErrorToast('Microsoft tokens are not available. Please log in.');
            return;
        }


      const response = await fetch(`${backendUrl}`, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${context.shortLivedToken}` // Use shortLivedToken
        },
        body: JSON.stringify({
          action: 'analyze_template',
          data: { templateId: selectedTemplate, boardId: context?.boardId },
          accessToken,
          refreshToken,
          expiresIn: (Number(tokenExpiry) - Date.now()) / 1000
        }),
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(`Backend error: ${response.status} - ${errorData.error || "Unknown error"}`);
      }

      const data = await response.json();
      if (data.success) {
        if (data.accessToken) {
            localStorage.setItem('msAccessToken', data.accessToken);
            localStorage.setItem('msRefreshToken', data.refreshToken);
            localStorage.setItem('msTokenExpiry', String(Date.now() + (data.expiresIn * 1000)));
        }
        setPlaceholders(data.placeholders);
        setColumns(data.columns);
        const initialMappings = {};
        data.placeholders.forEach((placeholder) => {
          initialMappings[placeholder.id] = '';
        });
        setPlaceholderMappings(initialMappings);
      } else {
        setError(data.error || 'Failed to analyze template');
        showErrorToast(data.error || 'Failed to analyze template');
      }
    } catch (e:any) {
      console.error(e);
      setError(e.message || 'Failed to analyze template.');
      showErrorToast(e.message || 'Failed to analyze template.');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    if (selectedTemplate && context?.boardId) {
      analyzeSelectedTemplate();
    }
  }, [selectedTemplate, context]);

  // --- Handle Placeholder Mapping Change ---
  const handlePlaceholderMappingChange = (placeholderId: string, columnId: string) => {
    setPlaceholderMappings({
      ...placeholderMappings,
      [placeholderId]: columnId,
    });
  };

  // --- Generate Document ---
const handleGenerateDocument = async () => {
    if (!selectedTemplate || !context?.itemId) {
        setError('Please select a template and ensure you are on an item.');
        showErrorToast('Please select a template and ensure you are on an item.');
        return;
    }

    setGenerating(true);
    setError('');

    try {
        const backendUrl = process.env.NEXT_PUBLIC_BACKEND_URL;
        if (!backendUrl) {
            throw new Error("NEXT_PUBLIC_BACKEND_URL is not defined.");
        }

        const accessToken = localStorage.getItem('msAccessToken');
        const refreshToken = localStorage.getItem('msRefreshToken');
        const tokenExpiry = localStorage.getItem('msTokenExpiry');

        if (!accessToken || !refreshToken || !tokenExpiry)
        {
            setError('Microsoft tokens are not available. Please log in.');
            showErrorToast('Microsoft tokens are not available. Please log in.');
            return;
        }


        const response = await fetch(`${backendUrl}`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${context.shortLivedToken}` // Use shortLivedToken
            },
            body: JSON.stringify({
                action: 'generate_document',
                data: {
                    templateId: selectedTemplate,
                    itemId: context.itemId,
                    placeholderMappings: placeholderMappings,
                },
                accessToken,
                refreshToken,
                expiresIn: (Number(tokenExpiry) - Date.now()) / 1000,
            }),
        });

        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(`Backend error: ${response.status} - ${errorData.error || "Unknown error"}`);
        }

        const data = await response.json();

        if (data.success) {
            if (data.accessToken) {
              localStorage.setItem('msAccessToken', data.accessToken);
              localStorage.setItem('msRefreshToken', data.refreshToken);
              localStorage.setItem('msTokenExpiry', String(Date.now() + (data.expiresIn * 1000)));
        }
            toast.success("Document generated and added to item files");
            monday.execute('notice', {
                message: `Document "${data.documentName}" generated successfully!`,
                type: 'success',
                timeout: 10000,
            });
            monday.execute('refresh');
        } else {
            setError(data.error || 'Failed to generate document');
            showErrorToast(data.error || 'Failed to generate document');
        }
    } catch (e:any) {
        console.error(e);
        setError(e.message || 'Failed to generate document.');
        showErrorToast(e.message || 'Failed to generate document.');
    } finally {
        setGenerating(false);
    }
};
  // --- Helper Function to Show Error Toasts ---
  const showErrorToast = (message: string) => {
    toast.error(message);
  };

  return (
    <div className={styles.container}>
      <Card>
        <CardHeader>
          <CardTitle className={styles.cardTitle}>Word Document Generator</CardTitle>
        </CardHeader>
        <CardContent>
          {error && <p className={styles.error}>{error}</p>}

          <Button onClick={handleMicrosoftLogin} disabled={loading} className={styles.button}>
            {loading ? 'Connecting...' : 'Connect to Microsoft OneDrive'}
          </Button>

          <Button onClick={loadTemplates} disabled={loadingTemplates} className={styles.button}>
            {loadingTemplates ? 'Loading Templates...' : 'Load Templates'}
          </Button>

          {templates.length > 0 && (
            <div className={styles.selectContainer}>
              <Select onValueChange={(value) => setSelectedTemplate(value)} value={selectedTemplate ?? undefined}>
                <SelectTrigger className={styles.selectTrigger}>
                  <SelectValue placeholder="Select a template" />
                </SelectTrigger>
                <SelectContent>
                  {templates.map((template) => (
                    <SelectItem key={template.id} value={template.id}>
                      {template.name}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>
          )}

          {placeholders.length > 0 && (
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead className={styles.tableHead}>Placeholder</TableHead>
                  <TableHead className={styles.tableHead}>Map to Column</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {placeholders.map((placeholder) => (
                  <TableRow key={placeholder.id}>
                    <TableCell className={styles.tableCell}>{placeholder.name}</TableCell>
                    <TableCell className={styles.tableCell}>
                      <Select
                        value={placeholderMappings[placeholder.id] || ''}
                        onValueChange={(value) => handlePlaceholderMappingChange(placeholder.id, value)}
                      >
                        <SelectTrigger>
                          <SelectValue placeholder="Select a column" />
                        </SelectTrigger>
                        <SelectContent>
                            <SelectItem key="name" value="name">Item Name</SelectItem>
                          {columns.map((column) => (
                            <SelectItem key={column.id} value={column.id}>
                              {column.title}
                            </SelectItem>
                          ))}
                        </SelectContent>
                      </Select>
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          )}

          <Button onClick={handleGenerateDocument} disabled={generating || loading} className={styles.button}>
            {generating ? 'Generating...' : 'Generate Document'}
          </Button>
        </CardContent>
      </Card>
      <Toaster /> {/* Add the Toaster component here */}
    </div>
  );
}