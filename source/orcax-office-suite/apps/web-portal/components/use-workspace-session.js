'use client';

import { useCallback, useEffect, useState } from 'react';
import { api } from '../lib/api';
import { clearAccessToken, writeAccessToken } from '../lib/session';

export function useWorkspaceSession() {
  const [session, setSession] = useState(null);
  const [workspaces, setWorkspaces] = useState([]);
  const [error, setError] = useState('');
  const [ready, setReady] = useState(false);

  const reload = useCallback(async () => {
    try {
      const me = await api('/auth/me');
      const list = await api('/workspaces');
      setSession(me);
      setWorkspaces(list.workspaces ?? []);
      setError('');
      return me;
    } catch (reason) {
      if (reason.status === 401) {
        clearAccessToken();
        setSession(null);
        setWorkspaces([]);
        setError('');
      } else {
        setError(reason.message);
      }
      return null;
    } finally {
      setReady(true);
    }
  }, []);

  useEffect(() => {
    let ignore = false;
    api('/auth/me')
      .then(async (me) => {
        const list = await api('/workspaces');
        if (ignore) {
          return;
        }
        setSession(me);
        setWorkspaces(list.workspaces ?? []);
        setError('');
      })
      .catch((reason) => {
        if (ignore) {
          return;
        }
        if (reason.status === 401) {
          clearAccessToken();
          setSession(null);
          setWorkspaces([]);
          setError('');
          return;
        }
        setError(reason.message);
      })
      .finally(() => {
        if (!ignore) {
          setReady(true);
        }
      });
    return () => {
      ignore = true;
    };
  }, []);

  async function switchWorkspace(workspaceId) {
    const next = await api('/workspaces/switch', {
      method: 'POST',
      body: { workspaceId, confirmed: true },
    });
    writeAccessToken(next.accessToken);
    setSession(next);
    const list = await api('/workspaces');
    setWorkspaces(list.workspaces ?? []);
  }

  async function logout() {
    try {
      await api('/auth/logout', { method: 'POST' });
    } finally {
      clearAccessToken();
      setSession(null);
      setWorkspaces([]);
    }
  }

  return { session, workspaces, error, ready, reload, switchWorkspace, logout, setSession };
}
