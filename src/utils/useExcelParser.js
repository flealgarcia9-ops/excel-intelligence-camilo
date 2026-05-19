import { useState, useCallback } from 'react';
import { parseExcelFile } from './excelParser';

/**
 * Hook para parsear Excel/CSV con progreso real.
 */
export function useExcelParser() {
  const [loading, setLoading] = useState(false);
  const [progress, setProgress] = useState(0);

  const parseFile = useCallback((file) => {
    return new Promise((resolve, reject) => {
      setLoading(true);
      setProgress(10);

      // Use setTimeout to allow UI to update before heavy parsing
      setTimeout(() => {
        setProgress(30);

        parseExcelFile(file)
          .then((result) => {
            setProgress(80);
            setTimeout(() => {
              setProgress(100);
              setTimeout(() => {
                setLoading(false);
                setProgress(0);
              }, 200);
              resolve(result);
            }, 100);
          })
          .catch((err) => {
            setLoading(false);
            setProgress(0);
            reject(err);
          });
      }, 50);
    });
  }, []);

  return { parseFile, loading, progress };
}
