/**
 * @OnlyCurrentDoc
 */

/** refresh guilds data */
function refreshData(): void {

  Core.refreshData();

}

/** workaround to tslint issue of namespace scope after importingtype definitions */
declare namespace SwgohHelp {

  function getGuildData(allycode: number): GuildData;
  function getPlayerData(allyCode: number): PlayerData;
  function getUnitList(): UnitsDefinitions;

}

/** Shortcuts for Google Apps Script classes */
const SPREADSHEET = SpreadsheetApp.getActive();

type URLFetchRequestOptions = GoogleAppsScript.URL_Fetch.URLFetchRequestOptions;

type KeyedType<T> = {
  [key: string]: T;
};

type KeyedNumbers = KeyedType<number>;

interface GuildData {
  id: number;
  name: string;
  members: PlayerData[];
}

interface PlayerData {
  allyCode: number;
  level?: number;
  link?: string;
  name: string;
  gp: number;
  heroesGp: number;
  shipsGp: number;
  fleetArenaBattlesWon: number;
  squadArenaBattlesWon: number;
  normalBattlesWon: number;
  hardBattlesWon: number;
  galacticWarBattlesWon: number;
  guildRaidsWon: number;
  guildTokensEarned: number;
  gearDonatedInGuildExchange: number;
  units: UnitInstances;
}

type UnitsDefinitions = {
  heroes: UnitDefinition[];
  ships: UnitDefinition[];
};

/** A unit's name, baseId and tags */
interface UnitDefinition {
  /** Unit Id */
  baseId: string;
  name: string;
  /** Alignment, role and tags */
  tags: string;
  type?: number;
}

type Ability = {
  name: string;
  tier: number;
  isZeta: boolean;
};

/**
 * A unit instance attributes
 * (baseId, gearLevel, level, name, power, rarity, stats, tags)
 */
interface UnitInstance {
  type: Units.TYPES;
  baseId?: string;
  gearLevel?: number;
  level: number;
  name?: string;
  power: number;
  abilities: Ability[];
  rarity: number;
  stats?: string;
  tags?: string;
}

type UnitInstances = {
  [key: string]: UnitInstance;
};

/** Constants for data source */
enum DATASOURCES {
  /** Use swgoh.gg API as data source */
  SWGOH_GG = 'SwgohGg',
  /** Use swgoh.help API as data source */
  SWGOH_HELP = 'SwgohHelp',
}

/** Constants for sheets name */
enum SHEETS {
  SETUP = 'coreSetup',
  MEMBERS = 'coreRoster',
  UNITS = 'coreUnits',
}

/** settings related functions */
namespace config {

  /** SwgohHelp related settings */
  export namespace SwgohHelp {

    /** Get the SwgohHelp API username */
    export function username(): string {

      const metaSWGOHLinkCol = 9;
      const metaSWGOHLinkRow = 15;
      const result = SPREADSHEET.getSheetByName(SHEETS.SETUP)
        .getRange(metaSWGOHLinkRow, metaSWGOHLinkCol)
        .getValue() as string;

      return result;
    }

    /** Get the SwgohHelp API password */
    export function password(): string {

      const metaSWGOHLinkCol = 9;
      const metaSWGOHLinkRow = 16;
      const result = SPREADSHEET.getSheetByName(SHEETS.SETUP)
        .getRange(metaSWGOHLinkRow, metaSWGOHLinkCol)
        .getValue() as string;

      return result;
    }

  }

}

namespace Units {

  export enum TYPES {
    HERO = 1,
    SHIP = 2,
  }

}

namespace Core {

  export const refreshData = (): void => {

    const activeGuilds = getActiveGuilds();
    const credentialIsFresh = isCredentialFresh();
    const rarIsFresh = isRARFresh();
    const guildsIsFresh = isGuildsFresh();

    let staleGuilds = guildsIsFresh && credentialIsFresh && rarIsFresh ? [] : activeGuilds;

    // let staleGuilds = activeGuilds.filter((e) => {

    //   if (rarIsFresh) {

    //     if (e.dataSource === DATASOURCES.SWGOH_GG) {
    //       return !isGuildsFresh();
    //     }
    //     if (e.dataSource === DATASOURCES.SWGOH_HELP) {
    //       return !(credentialIsFresh && isGuildFresh(e));
    //     }

    //   } else {  // TODO: intersect rar and guilds
    //     return true;
    //   }

    // });

    if (staleGuilds.length > 0) {

      staleGuilds = activeGuilds;  // TODO proper handling of stale

      const baseUnits = getBaseUnits();
      const datas = staleGuilds.map(e => getGuildData(e));
      renameAddRemove(staleGuilds, datas);
      setGuildNames(staleGuilds, datas);
      writeGuildNames(activeGuilds);
      writeRoster(datas);
      writeUnits(datas, baseUnits);
      refreshGuilds();
      refreshCredential();
      refreshRAR();
    } else {
      const UI = SpreadsheetApp.getUi();
      UI.alert(
        'Nothing to refresh',
        'Data from API is cached for one hour',
        UI.ButtonSet.OK,
      );

      return;
    }
  };

  const renameAddRemove = (settings: GuildSettings[], datas: GuildData[]): void => {
    const rar = getRAR();
    for (const add of rar.add) {
      const guild = add.guild;
      const target = datas.find(e => e.name === guild);
      if (target) {
        const allyCode = add.allyCode;
        // is in another guild?
        let found = false;
        for (const source of datas) {
          const i = source.members.findIndex(e => e.allyCode === allyCode);
          found = i > -1;
          if (found) {
            if (source !== target) {
              target.members.push(source.members.splice(i, 1)[0]);
            }
            break;
          }
        }
        if (!found) {
          // try to guess data source
          // const setup = settings.find(e => e.guildId === target.id);
          // const g = SwgohGg.getPlayerData(allyCode);
          // const h = SwgohHelp.getPlayerData(allyCode);
        }
      }
    }
    for (const rename of rar.rename) {
      const allyCode = rename.allyCode;
      for (const data of datas) {
        const i = data.members.findIndex(e => e.allyCode === allyCode);
        if (i > -1) {
          data.members[i].name = rename.name;
        }
      }
    }
    for (const allyCode of rar.remove) {
      for (const data of datas) {
        const i = data.members.findIndex(e => e.allyCode === allyCode);
        if (i > -1) {
          data.members.splice(i, 1);
        }
      }
    }
  };

  const setGuildNames = (settings: GuildSettings[], datas: GuildData[]): void => {
    for (const i in settings) {
      const setup = settings[i];
      if (setup.name.length === 0) {
        setup.name = datas[i].name;
      }
    }
    // TODO: duplicates guild name
  };

  const writeGuildNames = (settings: GuildSettings[]): void => {
    const values: string[][] = [];
    for (const setting of settings) {
      values.push([setting.name]);
    }
    SPREADSHEET.getSheetByName(SHEETS.SETUP)
      .getRange(3, 3, values.length)
      .setValues(values);
  };

  const writeRoster = (datas: GuildData[]): void => {
    const values: any[][] = [];
    for (const data of datas) {
      if (data) {
        const guildId = data.id;
        const guildName = data.name;
        for (const member of data.members) {
          values.push([
            guildId,
            guildName,
            member.name,
            member.allyCode,
            member.level,
            member.gp,
            member.heroesGp,
            member.shipsGp,
            member.fleetArenaBattlesWon,
            member.squadArenaBattlesWon,
            member.normalBattlesWon,
            member.hardBattlesWon,
            member.galacticWarBattlesWon,
            member.guildRaidsWon,
            member.guildTokensEarned,
            member.gearDonatedInGuildExchange,
          ]);
        }
      }
    }
    const headers = [
      'guildId',
      'guildName',
      'name',
      'allyCode',
      'level',
      'gp',
      'heroesGp',
      'shipsGp',
      'fleetArenaBattlesWon',
      'squadArenaBattlesWon',
      'normalBattlesWon',
      'hardBattlesWon',
      'galacticWarBattlesWon',
      'guildRaidsWon',
      'guildTokensEarned',
      'gearDonatedInGuildExchange',
    ];
    Sheets.setValues(
      SPREADSHEET.getSheetByName(SHEETS.MEMBERS),
      values,
      headers,
    );
  };

  const writeUnits = (datas: GuildData[], baseUnits: UnitDefinition[]): void => {
    const values: any[][] = [];
    let nAbilities = 0;
    let nColumns = 0;
    for (const data of datas) {
      if (data) {
        // const guildId = data.id;
        // const guildName = data.name;
        for (const member of data.members) {
          const units = member.units;
          for (const baseId in units) {
            const def = baseUnits.find(e => e.baseId === baseId);
            const unit = units[baseId];
            const columns = [
              // guildId,
              // guildName,
              member.name,
              member.allyCode,
              def.name,
              def.tags,
              unit.type,
              unit.rarity,
              unit.level,
              unit.gearLevel,
              unit.power,
            ];
            for (const ability of unit.abilities) {
              columns.push(ability.name);
              columns.push(ability.isZeta ? 'true' : 'false');
              columns.push(ability.tier);
            }
            nAbilities = Math.max(unit.abilities.length, nAbilities);
            nColumns = Math.max(columns.length, nColumns);
            values.push(columns);
          }
        }
      }
    }
    const headers = [
      // 'guildId',
      // 'guildName',
      'name',
      'allyCode',
      'unit',
      'tags',
      'type',
      'rarity',
      'level',
      'gearLevel',
      'power',
    ];
    for (let i = 0; i < nAbilities; i += 1) {
      headers.push(`ability${i}`);
      headers.push(`isZeta${i}`);
      headers.push(`tier${i}`);
    }
    Sheets.setValues(
      SPREADSHEET.getSheetByName(SHEETS.UNITS),
      values.map(e =>
        e.length !== nColumns ? e.concat(Array(nColumns).fill(null)).slice(0, nColumns) : e),
      headers,
    );
  };

  type GuildSettings = {
    name: string,
    dataSource: DATASOURCES,
    guildId?: number,
    allyCode?: number,
  };

  /** returns an array of active guild settings */
  const getActiveGuilds = (): GuildSettings[] => {

    const guilds = SPREADSHEET.getSheetByName(SHEETS.SETUP)
      .getRange(3, 2, 10, 5).getValues().reduce(
      (acc: GuildSettings[], columns: string[]) => {

        if (columns[0] === 'ON') {  // hardcoded checkbox value
          const name = columns[1];
          const dataSource = columns[2] as DATASOURCES;
          if (dataSource === DATASOURCES.SWGOH_GG) {
            acc.push({ name, dataSource, guildId: +columns[3] });
          } else if (dataSource === DATASOURCES.SWGOH_HELP) {
            acc.push({ name, dataSource, allyCode: +columns[4] });
          }
        }

        return acc;
      },
      [],
    ) as GuildSettings[];

    return guilds;
  };

  const getBaseUnits = (): UnitDefinition[] => {
    return SwgohGg.getHeroList().concat(SwgohGg.getShipList());
  };

  const getGuildData = (guild: GuildSettings): GuildData => {

    // Figure out which data source to use
    if (guild.dataSource === DATASOURCES.SWGOH_GG) {
      const data = SwgohGg.getGuildData(guild.guildId);
      // fix member level
      for (const member of data.members) {
        let level = 1;
        for (const baseId in member.units) {
          level = Math.max(member.units[baseId].level, level);
        }
        member.level = level;
      }
      return data;
    }
    if (guild.dataSource === DATASOURCES.SWGOH_HELP) {
      return SwgohHelp.getGuildData(guild.allyCode);
    }

  };

  /** check if guilds setup is fresh */
  const isGuildsFresh = (): boolean => {
    return getGuildsHash() === Cache.getGuildsHash();
  };

  /** updates cached guilds setup  */
  const refreshGuilds = (): void => {
    Cache.setGuildsHash(getGuildsHash());
  };

  /** compute hash for SwgohHelp guilds settings */
  const getGuildsHash = (): string => {

    const guilds = SPREADSHEET.getSheetByName(SHEETS.SETUP)
      .getRange(3, 2, 10, 5).getValues();

    return String(Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      JSON.stringify(guilds),
    ));
  };

  /** check if cached credential hash is fresh */
  const isCredentialFresh = (): boolean => {
    return getCredentialHash() === Cache.getCredentialHash();
  };

  /** updates cached credential hash */
  const refreshCredential = (): void => {
    Cache.setCredentialHash(getCredentialHash());
  };

  /** compute hash for SwgohHelp credential settings */
  const getCredentialHash = (): string => {

    const credential = SPREADSHEET.getSheetByName(SHEETS.SETUP)
      .getRange(15, 9, 2).getValues();

    return String(Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      JSON.stringify(credential),
    ));
  };

  /** check if cached rename/add/remove hash is fresh */
  const isRARFresh = (): boolean => {
    return getRARHash() === Cache.getRARHash();
  };

  /** updates cached rename/add/remove hash */
  const refreshRAR = (): void => {
    Cache.setRARHash(getRARHash());
  };

  type rarColumns = [string, string, number, number];

  /** compute hash for Rename/Add/Remove settings */
  const getRARHash = (): string => {

    const sheet = SPREADSHEET.getSheetByName(SHEETS.SETUP);

    const rar = sheet
      .getRange(15, 3, sheet.getMaxRows(), 4).getValues()
      .reduce(
        (acc: rarColumns[], e: rarColumns) => {
          if (e[2] > 0 || e[3] > 0) acc.push(e);
          return acc;
        },
        [],
      )
      .sort((a, b) => a[2] !== b[2] ? a[2] - b[2] : a[3] - b[3]) as rarColumns[];

    return String(Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      JSON.stringify(rar),
    ));
  };

  type RenameAddRemove = {
    rename: { name: string, allyCode: number }[],
    add: { guild: string, allyCode: number }[],
    remove: number[],
  };

  /** compute hash for Rename/Add/Remove settings */
  const getRAR = (): RenameAddRemove => {

    const sheet = SPREADSHEET.getSheetByName(SHEETS.SETUP);

    const rar = sheet
      .getRange(15, 3, sheet.getMaxRows(), 4).getValues()
      .reduce(
        (acc: RenameAddRemove, e) => {
          const guild = `${e[0]}`.trim();
          const name = `${e[1]}`.trim();
          const add = +e[2];
          const remove = +e[3];
          if (add > 0) {
            if (name.length > 0) {
              acc.rename.push({ name, allyCode: add });
            }
            if (guild.length > 0) {
              acc.add.push({ guild, allyCode: add });
            }
          }
          if (remove > 0) {
            acc.remove.push(remove);
          }
          return acc;
        },
        { rename: [], add: [], remove: [] },
      ) as RenameAddRemove;

    rar.remove = rar.remove.unique();

    return rar;
  };

  /** all things cached */
  namespace Cache {

    const cacheKey = SPREADSHEET.getId();
    const seconds = 3600;  // 1 hour
    const service = CacheService.getScriptCache();

    export const getGuildsHash = (): string =>
      service.get(`${cacheKey}-guilds`);

    export const setGuildsHash = (hash: string): void =>
      service.put(`${cacheKey}-guilds`, hash, seconds);

    export const getCredentialHash = (): string =>
      service.get(`${cacheKey}-cred`);

    export const setCredentialHash = (hash: string): void =>
      service.put(`${cacheKey}-cred`, hash, seconds);

    export const getRARHash = (): string =>
      service.get(`${cacheKey}-rar`);

    export const setRARHash = (hash: string): void =>
      service.put(`${cacheKey}-rar`, hash, seconds);

  }

  namespace Sheets {

    export const setValues = (sheet: GoogleAppsScript.Spreadsheet.Sheet, values, headers) => {

      values.unshift(headers);

      sheet.getRange(1, 1, values.length, headers.length)
        .setValues(values);

      const lastRow = sheet.getLastRow();
      const maxRows = sheet.getMaxRows();
      if (maxRows > lastRow) {
        sheet.deleteRows(lastRow + 1, maxRows - lastRow);
      }

      const lastCol = sheet.getLastRow();
      const maxcols = sheet.getMaxRows();
      if (maxcols > lastCol) {
        sheet.deleteColumns(lastCol + 1, maxcols - lastCol);
      }
    };

    export const numberOfCells = () => {

      const formatThousandsNoRounding = (n: number, dp?: number) => {
        const s = `${n}`;
        const b = n < 0 ? 1 : 0;
        const i = s.lastIndexOf('.');
        let j = i === -1 ? s.length : i;
        let r = '';
        const d = s.substr(j + 1, dp);
        while ((j -= 3) > b) {
          r = `,${s.substr(j, 3)}${r}`;
        }

        // tslint:disable-next-line:max-line-length
        return `${s.substr(0, j + 3)}${r}${dp ? `.${d}${d.length < dp ? ('00000').substr(0, dp - d.length) : ''}` : ''}`;
      };

      const sheets = SPREADSHEET.getSheets();
      let cellsCount = 0;
      for (const sheet of sheets) {
        cellsCount += (sheet.getMaxColumns() * sheet.getMaxRows());
      }
      // const ratio = cellsCount / 2000000 * 100;
      Logger.log(formatThousandsNoRounding(cellsCount));
    };

  }

}
