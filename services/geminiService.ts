
import { GradingReport, ClassData } from "../types";

export async function processClassGrading(data: ClassData): Promise<GradingReport> {
  const response = await fetch('/api/grade', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(data)
  });

  if (!response.ok) {
    const text = await response.text().catch(() => '');
    let errorMessage = '';
    try {
      const errorData = JSON.parse(text);
      errorMessage = errorData.error;
    } catch {
      if (text.includes('<pre>')) {
        const match = text.match(/<pre>([\s\S]*?)<\/pre>/);
        errorMessage = match ? match[1].trim() : text.substring(0, 500);
      } else {
        errorMessage = text ? text.substring(0, 500) : `HTTP Error ${response.status}: ${response.statusText}`;
      }
    }
    throw new Error(errorMessage || 'Failed to grade class');
  }

  return await response.json();
}

