import { useEffect, useState } from "react";

interface UserData {
    display_name: string;
    email: string;
}

export const useUserData = () => {
    const [user, setUser] = useState<UserData | null>(null);
    const [loading, setLoading] = useState(true);

    const token = localStorage.getItem("token");
    const refreshToken = localStorage.getItem("refresh_token");

    if (!token) {
        console.warn("No token found in localStorage");
        setLoading(false);
        return { user: null, loading: false };
    }



    useEffect(() => {
        const fetchUser = async () => {
            try {
                const myHeaders = new Headers();
                myHeaders.append("Content-Type", "application/json");
                myHeaders.append("Authorization", `Bearer ${token}`);
                if (refreshToken) myHeaders.append("x-refresh-token", refreshToken);



                const res = await fetch("https://app.right-report.com/api/addon-user", {
                    method: "POST",
                    headers: myHeaders,
                    redirect: "follow",
                });

                if (!res.ok) {
                    throw new Error(`HTTP error! status: ${res.status}`);
                }

                const data = await res.json();

                // The API returns the user under `data.user.user_metadata`
                const meta = data?.user?.user_metadata;
                setUser({
                    display_name: meta?.display_name || "Unknown User",
                    email: meta?.email || "No Email",
                });
            } catch (err) {
                console.error("Error fetching user data:", err);
            } finally {
                setLoading(false);
            }
        };

        fetchUser();
    }, []);

    return { user, loading };
};
