import type { FibRgFcLcb, FibRgLw, HeaderFooterRole } from '../types.js';
export type StoryKind = 'main' | 'footnote' | 'header' | 'macro' | 'comment' | 'endnote' | 'textbox' | 'headerTextbox';
export interface StoryWindow {
    kind: StoryKind;
    cpStart: number;
    cpEnd: number;
    length: number;
}
export interface PlcEntry {
    index: number;
    cpStart: number;
    cpEnd: number;
    data: Uint8Array;
}
export interface CommentRefMeta {
    initials: string;
    authorIndex: number;
    bookmarkId: number;
}
export interface TextboxMeta {
    reusable: boolean;
    shapeId: number;
    reserved: number;
}
export declare function buildStoryWindows(fibRgLw: FibRgLw, totalTextLength: number): Record<StoryKind, StoryWindow>;
export declare function readFixedPlc(tableBytes: Uint8Array, fc: number | undefined, lcb: number | undefined, dataSize: number): PlcEntry[];
export declare function readCpArray(tableBytes: Uint8Array, fc: number | undefined, lcb: number | undefined): number[];
export declare function parseSttbfRMark(tableBytes: Uint8Array, fibRgFcLcb: FibRgFcLcb): string[];
export declare function parseXstArray(tableBytes: Uint8Array, fibRgFcLcb: FibRgFcLcb): string[];
export declare function parseCommentRefMeta(data: Uint8Array): CommentRefMeta;
export declare function parseTextboxMeta(data: Uint8Array): TextboxMeta;
export interface HeaderStoryDescriptor {
    index: number;
    role: HeaderFooterRole;
    roleLabel: string;
    sectionIndex?: number;
    cpStart: number;
    cpEnd: number;
    inheritedFromSection?: number;
}
export declare function buildHeaderStoryDescriptors(tableBytes: Uint8Array, fibRgFcLcb: FibRgFcLcb, headerWindow: StoryWindow): HeaderStoryDescriptor[];
