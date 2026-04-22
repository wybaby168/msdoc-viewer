import type { DecodedProperty, FibRgFcLcb, SectionDescriptor, SectionPageSettings } from '../types.js';
export declare function sectionPropsToPageSettings(properties: DecodedProperty[]): SectionPageSettings;
export declare function readSections(wordBytes: Uint8Array, tableBytes: Uint8Array, fibRgFcLcb: FibRgFcLcb, mainStoryLength: number): SectionDescriptor[];
export declare function findSectionIndex(sections: SectionDescriptor[], cp: number): number;
