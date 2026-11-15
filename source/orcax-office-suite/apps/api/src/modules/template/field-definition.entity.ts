import { Column, Entity, Index, PrimaryColumn } from 'typeorm';

@Entity({ name: 'field_definitions' })
@Index('IDX_field_definitions_version_id', ['templateVersionId'])
export class FieldDefinition {
  @PrimaryColumn({ type: 'uuid' })
  id!: string;

  @Column({ name: 'template_version_id', type: 'uuid' })
  templateVersionId!: string;

  @Column({ name: 'field_key', type: 'varchar', length: 64 })
  fieldKey!: string;

  @Column({ type: 'varchar', length: 160 })
  label!: string;

  @Column({ name: 'field_type', type: 'varchar', length: 32 })
  fieldType!: string;

  @Column({ type: 'boolean' })
  required!: boolean;

  @Column({ name: 'default_value', type: 'text', nullable: true })
  defaultValue: string | null = null;

  @Column({ name: 'validation_json', type: 'jsonb' })
  validationJson!: object;

  @Column({ name: 'sort_order', type: 'int' })
  sortOrder!: number;
}
