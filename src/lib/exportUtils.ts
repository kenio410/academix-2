/**
 * Utility to save a file to the device or share it if supported.
 * This handles the differences between desktop browsers and mobile devices.
 */
export async function saveAndShareFile(
  fileName: string,
  data: string | Blob | ArrayBuffer,
  mimeType: string,
  options: { forceDownload?: boolean } = {}
): Promise<void> {
  try {
    let blob: Blob;
    
    if (data instanceof Blob) {
      blob = data;
    } else if (data instanceof ArrayBuffer) {
      blob = new Blob([data], { type: mimeType });
    } else {
      // Handle base64 string
      const byteCharacters = atob(data);
      const byteNumbers = new Uint8Array(byteCharacters.length);
      for (let i = 0; i < byteCharacters.length; i++) {
        byteNumbers[i] = byteCharacters.charCodeAt(i);
      }
      blob = new Blob([byteNumbers], { type: mimeType });
    }

    // Prepare filename - ensure it has the right extension and no weird characters for sharing
    const safeName = fileName.replace(/\s+/g, '_');

    // 2. Try Web Share API (best for mobile/tablets)
    if (!options.forceDownload && navigator.share && navigator.canShare) {
      const file = new File([blob], safeName, { type: mimeType });
      
      const shareData: ShareData = {
        files: [file],
        title: safeName,
      };

      if (navigator.canShare(shareData)) {
        try {
          await navigator.share(shareData);
          return; // Successfully shared
        } catch (shareError) {
          // If it's a cancelation, we might want to just stop
          if (shareError instanceof Error && shareError.name === 'AbortError') {
            return;
          }
          console.warn('Share failed, falling back to download:', shareError);
        }
      }
    }

    // 3. Fallback: Traditional Download (Standard for Desktop)
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = safeName;
    
    document.body.appendChild(link);
    link.click();
    
    setTimeout(() => {
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    }, 100);
    
  } catch (error) {
    console.error('Error saving/sharing file:', error);
    throw error;
  }
}
