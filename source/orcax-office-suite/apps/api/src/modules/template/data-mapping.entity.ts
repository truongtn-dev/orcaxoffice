import { Check, Column, Entity, Index, PrimaryColumn } from 'typeorm';

@Entity({ name: 'data_mappings' })
@Index('IDX_data_mappings_version_id', ['templateVersionId'])
@Check(`"source_kind" IN ('form', 'entity', 'constant')`)
@Check(`"transform" IN ('none', 'trim', 'upper', 'lower', 'date_iso')`)
export class DataMapping {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ name: 'template_version_id', type: 'uuid' })
  templateVersionId!: string;

  @Column({ name: 'field_key', type: 'varchar', length: 64 })
  fieldKey!: string;

  @Column({ name: 'source_kind', type: 'varchar', length: 32 })
  sourceKind!: 'form' | 'entity' | 'constant';

  @Column({ name: 'source_path', type: 'varchar', length: 200 })
  sourcePath!: string;

  @Column({ type: 'varchar', length: 32 })
  transform!: 'none' | 'trim' | 'upper' | 'lower' | 'date_iso';
}
