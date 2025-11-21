// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

export const apiVersion = "7.2-preview.1";
export const batchApiVersion = "5.0";
export const markdownCommentsApiVersion = "7.2-preview.4";

export function createEnumMapping<T extends Record<string, string | number>>(enumObject: T): Record<string, T[keyof T]> {
  const mapping: Record<string, T[keyof T]> = {};
  for (const [key, value] of Object.entries(enumObject)) {
    if (typeof key === "string" && typeof value === "number") {
      mapping[key.toLowerCase()] = value as T[keyof T];
    }
  }
  return mapping;
}

export function mapStringToEnum<T extends Record<string, string | number>>(value: string | undefined, enumObject: T, defaultValue?: T[keyof T]): T[keyof T] | undefined {
  if (!value) return defaultValue;
  const enumMapping = createEnumMapping(enumObject);
  return enumMapping[value.toLowerCase()] ?? defaultValue;
}

/**
 * Maps an array of strings to an array of enum values, filtering out invalid values.
 * @param values Array of string values to map
 * @param enumObject The enum object to map to
 * @returns Array of valid enum values
 */
export function mapStringArrayToEnum<T extends Record<string, string | number>>(values: string[] | undefined, enumObject: T): T[keyof T][] {
  if (!values) return [];
  return values.map((value) => mapStringToEnum(value, enumObject)).filter((v): v is T[keyof T] => v !== undefined);
}

/**
 * Converts a TypeScript numeric enum to an array of string keys for use with z.enum().
 * This ensures that enum schemas generate string values rather than numeric values.
 * @param enumObject The TypeScript enum object
 * @returns Array of string keys from the enum
 */
export function getEnumKeys<T extends Record<string, string | number>>(enumObject: T): string[] {
  return Object.keys(enumObject).filter((key) => isNaN(Number(key)));
}

/**
 * Safely converts a string enum key to its corresponding enum value.
 * Validates that the key exists in the enum before conversion.
 * @param enumObject The TypeScript enum object
 * @param key The string key to convert
 * @returns The enum value if key is valid, undefined otherwise
 */
export function safeEnumConvert<T extends Record<string, string | number>>(enumObject: T, key: string | undefined): T[keyof T] | undefined {
  if (!key) return undefined;

  const validKeys = getEnumKeys(enumObject);
  if (!validKeys.includes(key)) {
    return undefined;
  }

  return enumObject[key as keyof T];
}

/**
 * Encodes `>` and `<` for Markdown formatted fields.
 *
 * @param value The text value to encode
 * @param format The format of the field ('Markdown' or 'Html')
 * @returns The encoded text, or original text if format is not Markdown
 */
export function encodeFormattedValue(value: string, format?: "Markdown" | "Html"): string {
  if (!value || format !== "Markdown") return value;
  const result = value.replace(/</g, "&lt;").replace(/>/g, "&gt;");
  return result;
}

export function filterJsonByPaths(data: any, propertyFilters?: string[], currentPath: string = ''): any {
  // If no filters provided, return original data
  if (!propertyFilters || propertyFilters.length === 0) {
    return data;
  }

  // If data is not an object, return as is
  if (data === null || typeof data !== 'object') {
    return data;
  }

  // Handle arrays - apply function for each element
  if (Array.isArray(data)) {
    return data.map(item => filterJsonByPaths(item, propertyFilters, currentPath));
  }

  // Handle objects
  const result: any = {};

  for (const key in data) {
    if (!data.hasOwnProperty(key)) {
      continue;
    }

    // Build the current property path
    const propertyPath = currentPath ? `${currentPath}.${key}` : key;

    // Check if current path matches any filter
    let hasExactMatch = false;
    let hasPrefix = false;

    for (const filter of propertyFilters) {
      if (filter === propertyPath) {
        // Exact match - copy the entire property to result
        hasExactMatch = true;
        break;
      } else if (filter.startsWith(propertyPath + '.')) {
        // Filter starts with current path - need to recurse
        hasPrefix = true;
      }
    }

    if (hasExactMatch) {
      // Copy the entire property value
      result[key] = data[key];
    } else if (hasPrefix) {
      // Recursively filter the nested value
      const filtered = filterJsonByPaths(data[key], propertyFilters, propertyPath);

      // Only add if the filtered result has content
      if (filtered !== null && typeof filtered === 'object') {
        if (Array.isArray(filtered)) {
          if (filtered.length > 0) {
            result[key] = filtered;
          }
        } else if (Object.keys(filtered).length > 0) {
          result[key] = filtered;
        }
      }
    }
    // If no match and no prefix - skip this property
  }

  return result;
}
