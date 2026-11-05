import * as React from "react";
import {
  getMatterDriveId,
  listFolderContents,
  listMatters,
  type IMatter,
  type IMatterFile,
  type IMatterFolder,
} from "../services/matterService";
import { friendlyAuthError } from "../auth/authErrors";

export const ROOT_FOLDER = "root";

export interface IMatterLocation {
  matter: IMatter;
  driveId: string;
  /** Folders from the library root down to the open folder ([] = library root). */
  folderPath: IMatterFolder[];
}

export interface IMatterFolders {
  state: {
    matters: IMatter[] | undefined;
    location: IMatterLocation | undefined;
    folders: IMatterFolder[] | undefined;
    files: IMatterFile[] | undefined;
    isLoading: boolean;
    error: string | undefined;
  };
  actions: {
    loadMatters: () => Promise<IMatter[] | undefined>;
    selectMatter: (matter: IMatter) => Promise<boolean>;
    /** Opens a known location (pre-fill); false when it cannot be read. */
    openLocation: (location: IMatterLocation) => Promise<boolean>;
    openFolder: (folder: IMatterFolder) => Promise<void>;
    openPathLevel: (depth: number) => Promise<void>;
    clearMatter: () => void;
    setError: (error: string | undefined) => void;
  };
  computed: {
    folderId: string;
    locationLabel: string;
  };
}

/**
 * Matter list + library browsing for the Outlook panes (Matter → Folder steps). Only the newest navigation's result
 * is applied, so a slow answer for a matter the user already left never shows (or becomes the target).
 */
export function useMatterFolders(): IMatterFolders {
  const [matters, setMatters] = React.useState<IMatter[] | undefined>(undefined);
  const [location, setLocation] = React.useState<IMatterLocation | undefined>(undefined);
  const [folders, setFolders] = React.useState<IMatterFolder[] | undefined>(undefined);
  const [files, setFiles] = React.useState<IMatterFile[] | undefined>(undefined);
  const [isLoading, setIsLoading] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const latest = React.useRef(0);

  /** Opens `target` (its contents first, then the location); false when superseded or failed. */
  const navigate = React.useCallback(
    async (resolve: () => Promise<IMatterLocation>): Promise<boolean> => {
      const ticket = ++latest.current;
      setIsLoading(true);
      setError(undefined);
      try {
        const target = await resolve();
        const open = target.folderPath[target.folderPath.length - 1]?.id ?? ROOT_FOLDER;
        const contents = await listFolderContents(target.driveId, open);
        if (ticket !== latest.current) return false;
        setLocation(target);
        setFolders(contents.folders);
        setFiles(contents.files);
        return true;
      } catch (e) {
        if (ticket === latest.current) setError(friendlyAuthError(e));
        return false;
      } finally {
        if (ticket === latest.current) setIsLoading(false);
      }
    },
    []
  );

  const loadMatters = React.useCallback(async () => {
    setIsLoading(true);
    setError(undefined);
    try {
      const list = await listMatters();
      setMatters(list);
      return list;
    } catch (e) {
      setError(friendlyAuthError(e));
      return undefined;
    } finally {
      setIsLoading(false);
    }
  }, []);

  const selectMatter = React.useCallback(
    (matter: IMatter) =>
      navigate(async () => ({
        matter,
        driveId: await getMatterDriveId(matter.siteId),
        folderPath: [],
      })),
    [navigate]
  );

  const openLocation = React.useCallback(
    (target: IMatterLocation) => navigate(async () => target),
    [navigate]
  );

  const openFolder = React.useCallback(
    async (folder: IMatterFolder) => {
      if (location)
        await navigate(async () => ({ ...location, folderPath: [...location.folderPath, folder] }));
    },
    [location, navigate]
  );

  const openPathLevel = React.useCallback(
    async (depth: number) => {
      if (location)
        await navigate(async () => ({
          ...location,
          folderPath: location.folderPath.slice(0, depth),
        }));
    },
    [location, navigate]
  );

  const clearMatter = React.useCallback(() => {
    latest.current++;
    setIsLoading(false);
    setError(undefined);
    setLocation(undefined);
    setFolders(undefined);
    setFiles(undefined);
  }, []);

  const folderId = location?.folderPath[location.folderPath.length - 1]?.id ?? ROOT_FOLDER;
  const locationLabel = React.useMemo(
    () =>
      location ? [location.matter.ref, ...location.folderPath.map((f) => f.name)].join(" › ") : "",
    [location]
  );

  return {
    state: { matters, location, folders, files, isLoading, error },
    actions: {
      loadMatters,
      selectMatter,
      openLocation,
      openFolder,
      openPathLevel,
      clearMatter,
      setError,
    },
    computed: { folderId, locationLabel },
  };
}
