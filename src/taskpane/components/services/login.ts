
export interface LoginResponse {
  error?: string;
  token?: string;
  message?: string;
  [key: string]: any; 
}

export async function loginUser(email: string, password: string): Promise<LoginResponse> {
  try {
    const response = await fetch("https://app.right-report.com/api/addon-login", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ email, password }),
    });

    let result: LoginResponse | null = null;

    try {
      result = await response.json();
    } catch {
      result = null;
    }

    if (result?.error) {
      // Wrong credentials
      return { error: result.error };
    }

    // Success
    return result || {};
  } catch (error) {
    console.error("API Error:", error);
    return { error: "Network error. Please try again." };
  }
}
